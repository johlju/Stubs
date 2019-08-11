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
        Get-Command -Module ActiveDirectory | New-StubCmdlet

        Create stubs of all the commands in the ActiveDirectory module.

    .NOTES
        Change log:
            2019-08-11 - Johan Ljunggren - Created.
#>
function New-StubCmdlet
{
    # Suppressed because this command does not change state.
    [System.Diagnostics.CodeAnalysis.SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding(DefaultParameterSetName = 'CommandInfo')]
    [OutputType([System.String])]
    param
    (
        [Parameter(Position = 0, Mandatory = $true, ParameterSetName = 'CommandName')]
        [System.String]
        $CommandName,

        [Parameter(ValueFromPipeline = $true, ParameterSetName = 'CommandInfo')]
        [System.Management.Automation.CommandInfo]
        $CommandInfo,

        [Parameter()]
        [ValidateScript( {
                if ($null -ne $_.Ast.ParamBlock -or $null -ne $_.Ast.DynamicParamBlock)
                {
                    throw (New-Object ArgumentException ("The function body scriptblock cannot contain Param or DynamicParam blocks"))
                }
                else
                {
                    $true
                }
            })]
        [System.Management.Automation.ScriptBlock]
        $FunctionBody,

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
                [System.Management.Automation.PSCmdlet]::CommonParameters
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

                [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
                    WhatIf
                    Confirm
                    UseTransaction
            #>
            # $commonParameters = [System.String[]] [System.Management.Automation.PSCmdlet]::CommonParameters
            # $optionalCommonParameters = [System.String[]] [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
        }
    }

    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'CommandInfo')
        {
            try
            {
                $script = New-Object -TypeName System.Text.StringBuilder

                $null = $script.AppendFormat('function {0}', $CommandInfo.Name).AppendLine()
                $null = $script.AppendLine('{')

                # Write help
                $helpContent = Get-Help -Name $CommandInfo.Name -Full
                if ($helpContent.Synopsis)
                {
                    $null = $script.AppendLine('<#')
                    $null = $script.AppendLine('.SYNOPSIS')
                    $null = $script.AppendFormat('    {0}', $helpContent.Synopsis.Trim()).AppendLine()

                    foreach ($parameter in $CommandInfo.Parameters.Keys)
                    {
                        if (
                            $parameter -notin [System.Management.Automation.PSCmdlet]::CommonParameters `
                            -and $parameter -notin [System.Management.Automation.PSCmdlet]::OptionalCommonParameters)
                        {
                            $parameterHelp = ($helpContent.Parameters.Parameter |
                                Where-Object { $_.Name -eq $parameter }).Description.Text

                            if ($parameterHelp)
                            {
                                $paragraphs = $parameterHelp.Split("`n", [StringSplitOptions]::RemoveEmptyEntries)

                                $null = $script.AppendFormat('.PARAMETER {0}', $parameter).AppendLine()

                                foreach ($paragraph in $paragraphs)
                                {
                                    $null = $script.AppendFormat('    {0}', $paragraph).AppendLine()
                                }
                            }
                        }
                    }

                    $null = $script.AppendLine('#>').AppendLine()
                }

                # Write CmdletBinding
                if ($cmdletBindingAttribute = [System.Management.Automation.ProxyCommand]::GetCmdletBindingAttribute($CommandInfo))
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

                    if ($param = [System.Management.Automation.ProxyCommand]::GetParamBlock($CommandInfo))
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
