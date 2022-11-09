# Using xlwings to synchrozine vba with Visual Studio Code

## Use conda prompt
In VS CODE:
  - Press F1 > Select Python Interpreter > Select conda python interpreter
  - Start a new CMD prompt, now it'll behave as a conda prompt.

In order to use xlwings synchronization must have an updated version beyond the 26.00 version

## Update xlwings
- `pip --upgrade xlwings`

## Using xlwings synchronization
Synchronization is used to edit the vba code in VS Codes while seeing changes in the excel file. To use it properly it's better to have a control which executes the code. Sometimes while having both editors opened synchronization doesn't work properly. 

To fix this you could write code only in VS Code and execute the code in the excel file using a control such as a button or menu.

- Get into the repository where you are going to store your VBA project, there should be stored the excel file with the project.

- Execute
`xlwings vba edit` 'To synchronize normal excel files (current opened file)
`xlwings vba edit -f filename` ''To synchronize .xlam files

## Exporting from VBA Project to local repository
Some people are used to work with the VBA editor, when you have made changes to the code and want to commit those changes:

- Get into the repository where you are going to store your VBA project
- Execute
`xlwings vba export filename` 'To export normal excel files (current opened file)
`xlwings vba export -f filename` ''To export .xlam files
