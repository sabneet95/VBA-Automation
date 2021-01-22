# VBA Automation - Workflow Pipelines

A repository of VBA-based automation solutions for atmospheric sciences and neurophysics.

![](https://github.com/sabneet95/VBA-Automation/blob/master/vba.jpg)

→ `Domain-specific code! Will not work on its own but can be modified for other projects.`

## Requirements

[VBA 7 or above](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)

## Build Tested

Microsoft Excel
* Version: 16.0.13714.20000 64-bit
* OS: Windows_NT x64 10.0.19042
* Memory: 1981M
* Cores: 8

## Usage

1)	Open the project in **Microsoft Excel** > under the Developer tab _run_ the VBA Macro as desired

```VBA

Function Col_lett(ByVal ColumnNumber As Integer)
Col_lett = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function
Sub Weather()
Dim Width As Single, Height As Single, NumWide As Long

.
..
...

```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.


## License
[MIT](https://choosealicense.com/licenses/mit/)
