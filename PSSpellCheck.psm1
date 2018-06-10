function Invoke-SpellCheck() {
    [CmdletBinding()]
    param (
        # Target file name to perform spell check
        [Parameter(Mandatory,ValueFromPipeline)]
        [System.IO.FileInfo] $Path
    )
    process {
        if(Test-Path $Path){
            add-type -AssemblyName "Microsoft.Office.Interop.Word" 
            $word = New-Object -comobject word.application        
            $SpecialChars = ',./<>?[]\{}|_+-=!@#$%^&*'.ToCharArray()
            $i=1
            
            (Get-Content $Path).ForEach({
                $_ -split [System.Environment]::NewLine
            }).ForEach({ # every line
                    $_.split(" ").trim().Where({
                        ![string]::IsNullOrWhiteSpace($_) -and 
                        $_ -notlike '*@*' -and 
                        $_ -notlike '*``*' -and 
                        $_ -notin $SpecialChars -and 
                        $_.length -gt 3
                    }).foreach({ # every word
                        if(!$word.CheckSpelling($_)){ # checkspelling() return $false           
                            [PSCustomObject]@{
                                LineNum = $i
                                Words = $_
                                Filename = $Path
                            }
                        }
                    })
                    $i++
            })
        } # end of test-path condition
        else{
            throw "$path does't exists!"
        }

    }
}

Export-ModuleMember *-*
