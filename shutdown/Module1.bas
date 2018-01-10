Attribute VB_Name = "Module1"
Public Type SYSTEM_POWER_STATUS
ACLineStatus  As Byte
BatteryFlag As Byte
BatteryLifePercent As Byte
Reservedl As Byte
BatteryLifeTime As Byte
BatteryyFullLifeTime As Long

End Type
Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

