(Get-ChildItem *.md -Recurse).FullName | 
Where-Object{$_ -like '*docs*README.md'} | 
ForEach-Object {
    Invoke-SpellCheck -Path $_
}
