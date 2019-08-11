<#
    .SYNOPSIS
        Create a new stub of a cmdlet.

    .DESCRIPTION
        New-StubCmdlet recreates a command as a function with param block and
        dynamic param block (if used).

    .PARAMETER CommandName
        Specifies the name of the cmdlet to generate a stub for.

    .PARAMETER CommandInfo
        Specifies a System.Management.Automation.CommandInfo object of the cmdlet
        to generate a stub for. This is normally acquired by the cmdlet Get-Command.

    .PARAMETER FunctionBody
        Specifies a script block of an optional function body to attach to the
        generated stub.

    .INPUTS
        System.Management.Automation.CommandInfo

    .OUTPUTS
        System.String

    .EXAMPLE
        New-StubCmdlet Test-Path

        Create a stub of the Test-Path command.

    .EXAMPLE
        Get-Command -Module AppLocker | New-StubCmdlet

        Create a stub of all commands in the AppLocker module.

    .NOTES
        Change log:
            2019-08-11 - Johan Ljunggren - Created.
#>
function New-StubCmdlet
{
    # Suppressed because this command does not change state.
    [System.Diagnostics.CodeAnalysis.SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding(DefaultParameterSetName = 'FromPipeline')]
    [OutputType([System.String])]
    param
    (
        [Parameter(Position = 0, Mandatory, ParameterSetName = 'CommandName')]
        [System.String]
        $CommandName,

        [Parameter(ValueFromPipeline, ParameterSetName = 'CommandInfo')]
        [System.Management.Automation.CommandInfo]
        $CommandInfo,

        [Parameter()]
        [ValidateScript( {
                if ($null -ne $_.Ast.ParamBlock -or $null -ne $_.Ast.DynamicParamBlock)
                {
                    throw (New-Object ArgumentException ("FunctionBody scriptblock cannot contain Param or DynamicParam blocks"))
                }
                else
                { $true
                }
            })]
        [scriptblock]$FunctionBody,

        [Parameter()]
        [System.Collections.Hashtable[]]
        $ReplaceTypeDefinition
    )

    begin
    {
        if ($PSCmdlet.ParameterSetName -eq 'CommandName')
        {
            $null = $PSBoundParameters.Remove('CommandName')
            Get-Command $CommandName | New-StubCmdlet @PSBoundParameters
        }
        else
        {
            <#
            PS C:\Users\johan.ljunggren> [System.Management.Automation.PSCmdlet]::CommonParameters
            Verbose
            Debug
            ErrorAction
            WarningAction
            InformationAction
            ErrorVariable
            WarningVariable
            OutVariable
            OutBuffer
            PipelineVariable
            InformationVariable
            PS C:\Users\johan.ljunggren> [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            WhatIf
            Confirm
            UseTransaction
            #>
            $commonParameters = ([CommonParameters]).GetProperties().Name
            $shouldProcessParameters = ([ShouldProcessParameters]).GetProperties().Name
        }
    }

    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'CommandInfo')
        {
            try
            {
                $script = New-Object -TypeName System.Text.StringBuilder

                $null = $script.AppendFormat('function {0} {{', $CommandInfo.Name).
                AppendLine()

                # Write help
                $helpContent = Get-Help $CommandInfo.Name -Full
                if ($helpContent.Synopsis)
                {
                    $null = $script.AppendLine('<#').
                    AppendLine('.SYNOPSIS').
                    AppendFormat('    {0}', $helpContent.Synopsis.Trim()).
                    AppendLine()

                    foreach ($parameter in $CommandInfo.Parameters.Keys)
                    {
                        if ($parameter -notin $commonParameters -and $parameter -notin $shouldProcessParameters)
                        {
                            $parameterHelp = ($helpcontent.parameters.parameter | Where-Object { $_.Name -eq $parameter }).Description.Text
                            if ($parameterHelp)
                            {
                                $paragraphs = $parameterHelp.Split("`n", [StringSplitOptions]::RemoveEmptyEntries)

                                $null = $script.AppendFormat('.PARAMETER {0}', $parameter).
                                AppendLine()

                                foreach ($paragraph in $paragraphs)
                                {
                                    $null = $script.AppendFormat('    {0}', $paragraph).
                                    AppendLine()
                                }
                            }
                        }
                    }
                    $null = $script.AppendLine('#>').
                    AppendLine()
                }

                # Write CmdletBinding
                if ($cmdletBindingAttribute = [ProxyCommand]::GetCmdletBindingAttribute($CommandInfo))
                {
                    $null = $script.AppendLine($cmdletBindingAttribute)
                }

                # Write OutputType
                foreach ($outputType in $CommandInfo.OutputType)
                {
                    $null = $script.Append('[OutputType(')
                    if ($outputType.Type)
                    {
                        $null = $script.AppendFormat('[{0}]', $outputType.Type)
                    }
                    else
                    {
                        $null = $script.AppendFormat("'{0}'", $outputType.Name)
                    }
                    $null = $script.AppendLine(')]')
                }

                # Write param
                if ($CommandInfo.CmdletBinding -or $CommandInfo.Parameters.Count -gt 0)
                {
                    $null = $script.Append('param (')

                    if ($param = [ProxyCommand]::GetParamBlock($CommandInfo))
                    {
                        foreach ($line in $param -split '\r?\n')
                        {
                            if ($PSBoundParameters.ContainsKey('ReplaceTypeDefinition'))
                            {
                                foreach ($type in $ReplaceTypeDefinition)
                                {
                                    if ($line -match ('\[{0}\]' -f $type.ReplaceType))
                                    {
                                        $line = $line -replace $type.ReplaceType, $type.WithType
                                    }
                                }
                            }

                            $null = $script.AppendLine($line.Trim())
                        }
                    }
                    else
                    {
                        $null = $script.Append(' ')
                    }

                    $null = $script.AppendLine(')')
                }

                $newStubDynamicParamArguments = @{
                    CommandInfo = $CommandInfo
                }

                if ($PSBoundParameters.ContainsKey('ReplaceTypeDefinition'))
                {
                    $newStubDynamicParamArguments['ReplaceTypeDefinition'] = $ReplaceTypeDefinition
                }

                if ($dynamicParams = New-StubDynamicParam @newStubDynamicParamArguments)
                {
                    # Write dynamic params
                    $null = $script.AppendScript($dynamicParams)
                }

                # Insert function body, if specified
                if ($null -ne $FunctionBody)
                {
                    if ($null -ne $FunctionBody.Ast.BeginBlock)
                    {
                        $null = $script.AppendLine(($FunctionBody.Ast.BeginBlock))
                    }

                    if ($null -ne $FunctionBody.Ast.ProcessBlock)
                    {
                        $null = $script.AppendLine(($FunctionBody.Ast.ProcessBlock))
                    }

                    if ($null -ne $FunctionBody.Ast.EndBlock)
                    {
                        if ($FunctionBody.Ast.EndBlock -imatch '\s*end\s*{')
                        {
                            $null = $script.AppendLine(($FunctionBody.Ast.EndBlock))
                        }
                        else
                        {
                            # Simple scriptblock does not explicitly specify that code is in end block, so we add the block decoration
                            $null = $script.AppendLine('end {')
                            $null = $script.AppendLine(($FunctionBody.Ast.EndBlock))
                            $null = $script.AppendLine('}')
                        }
                    }
                }

                # Close the function

                $null = $script.AppendLine('}')

                $script.ToString()
            }
            catch
            {
                Write-Error -ErrorRecord $_
            }
        }
    }
}
