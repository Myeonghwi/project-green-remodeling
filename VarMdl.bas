Attribute VB_Name = "VarMdl"
Option Explicit

Public Const BATCH_FILE_PATH            As String = "\energyplus\EnergyPlusV8-1-0\RunEPlus.bat TestCase3 KOR_Inchon.471120_IWEC"
Public Const PROGRESS_PATH              As String = "\energyplus\EnergyPlusV8-1-0\ProcessFiles\Progress\84m2\TestCase3.idf"
Public Const TEMPLATE_PATH              As String = "\energyplus\EnergyPlusV8-1-0\ProcessFiles\Template\84m2\TestCase3.idf"
Public Const OUTPUT_PATH                As String = "\energyplus\EnergyPlusV8-1-0\ProcessFiles\Outputs"
Public Const API_XML_PATH               As String = "\files\api\"

Public Const strTitle                   As String = "▒ Green Remodeling Decision Making System"
Public Const strCoName                  As String = "▒ www.yonsei.ac.kr"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''The Repla_Table''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const VAR_NAME                   As Integer = 1          '변수명
Public Const IS_SELECTION               As Integer = 2          '선택여부
Public Const IS_RANGE                   As Integer = 3          '범위여부
Public Const REPLA_VALUE                As Integer = 5          '설정값
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public rowCount                         As Integer
Public colCount                         As Integer

Public rngMax                           As Double
Public rngMin                           As Double
Public term                             As Double

Public montlyElecConsumption(12)        As Double
Public montlyGasConsumption(12)         As Double

Public lst()                            As Variant

Public blnCheck                         As Boolean
