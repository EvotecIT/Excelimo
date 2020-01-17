<p align="center">
  <a href="https://dev.azure.com/evotecpl/Excelimo/_build/latest?definitionId=3"><img src="https://dev.azure.com/evotecpl/Excelimo/_apis/build/status/EvotecIT.Excelimo"></a>
  <a href="https://www.powershellgallery.com/packages/Excelimo"><img src="https://img.shields.io/powershellgallery/v/Excelimo.svg"></a>
  <a href="https://www.powershellgallery.com/packages/Excelimo"><img src="https://img.shields.io/powershellgallery/vpre/Excelimo.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
  <a href="https://github.com/EvotecIT/Excelimo"><img src="https://img.shields.io/github/license/EvotecIT/Excelimo.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/Excelimo"><img src="https://img.shields.io/powershellgallery/p/Excelimo.svg"></a>
  <a href="https://github.com/EvotecIT/Excelimo"><img src="https://img.shields.io/github/languages/top/evotecit/Excelimo.svg"></a>
  <a href="https://github.com/EvotecIT/Excelimo"><img src="https://img.shields.io/github/languages/code-size/evotecit/Excelimo.svg"></a>
  <a href="https://www.powershellgallery.com/packages/Excelimo"><img src="https://img.shields.io/powershellgallery/dt/Excelimo.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

# Excelimo - PowerShell Module

This module allows to create Excel files from PowerShell using DSL syntax. Since I want to limit amount of modules that I have to support I decided to move code back to PSWriteExcel so that there's only one package for both PSWriteExcel and Excelimo. Any further development will happen there.

This project is archived. For development go to https://github.com/EvotecIT/PSWriteExcel

```PowerShell
Uninstall-Module Emailimo -AllVersions
Install-Module PSWriteHTML -Force
```

## Updates for archive purposes

- 0.0.4 / 2020.01.17
  - Moved to PSWriteExcel for better development and more features
  - This last release basically points PSD1 file to PSWriteExcel. You don't need it to use Excelimo anymore.

- 0.0.2 / 2019.06.23
  - Support for PSWriteExcel 0.1.2
  - Support for PSSharedGoods 0.0.79
