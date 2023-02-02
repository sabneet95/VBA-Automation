# VBA Automation for Atmospheric Sciences and Neurophysics

This repository contains VBA-based automation solutions designed specifically for the atmospheric sciences and neurophysics domains.

![](https://github.com/sabneet95/VBA-Automation/blob/master/vba.jpg)

→ `Please note that the code provided is domain-specific and may not work as-is for other projects, but it can be modified to suit your needs.`

## Requirements

[VBA 7 or higher](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)

## Build Tested

Microsoft Excel
* Version: 16.0.13714.20000 64-bit
* OS: Windows_NT x64 10.0.19042
* Memory: 1981M
* Cores: 8

## How to Use

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

Contributions are welcome. If you would like to make major changes, please open an issue to discuss the proposed changes before making a pull request.

Please ensure that appropriate tests are updated when making changes.


## License
This repository is licensed under the [MIT](https://choosealicense.com/licenses/mit/) License.
