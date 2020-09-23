Attribute VB_Name = "ModBus"
Option Explicit
'Author:Tom Pydeski
'BitWise Industrial Automation
'F & L Machinery Design, Inc.
'This project communicates with devices via modbus
'it was originally intended for ascii mode, but is
'being used to talk to Emerson servo drives, which
'communicate via ModBus RTU
'It was tested with an EN-204 drive.
'below is the list of parameters By Modbus Address from the
'Emerson Control Techniqes Reference Manual
'P/N 400504-01
'Revision: A5
'Date: October 18, 2002
'
'© BitWise Industrial Automation, 2004
'
'function codes
'1 - Read Coil Status
'2 - Read Input Status
'3 - Read Holding Registers
'4 - Read Input Registers
'5 - Force Single Coil
'6 - Preset Single Register
'7 - Read Exception Status
'15 - Force Multiple Coils
'16 - Preset Multiple Registers
'
'Modbus Name Type Units NVM Range DecimalPlaces Access Eb Ei EN FM1 FM2 MDS
'2 - 15 Input Line Force On/Off Command Array BIT N RW Eb Ei EN FM1 FM2 MDS
'18 - 31 Input Line Force On/Off Enable Array BIT N RW Eb Ei EN FM1 FM2 MDS
'33 - 48 Output Line Force On/Off Command Array BIT N RW Eb Ei EN FM1 FM2 MDS
'49 - 64 Output Line Force On/Off Enable Array BIT N RW Eb Ei EN FM1 FM2 MDS
'65 - 80 Output Line Active Off Array BIT N RW Eb Ei EN FM1 FM2 MDS
'129 - 160 Input Function Active Off Array BIT N RW Eb Ei EN FM1 FM2 MDS
'257 - 288 Input Function Always Active Array BIT N RW Eb Ei EN FM1 FM2 MDS
'1001 Warmstart Execute BIT N RW Eb Ei EN FM1 FM2 MDS
'1002 Write RAM to NVM BIT N RW Eb Ei EN FM1 FM2 MDS
'1003 Read NVM to RAM BIT N RW Eb Ei EN FM1 FM2 MDS
'1004 Update Predefined Setup BIT N RW Eb Ei EN FM1 FM2 MDS
'1007 Clear Fault BIT N RW Eb Ei EN FM1 FM2 MDS
'1101 - 1108 Motion Command Execute Array BIT N RW Ei FM2
'1151 Stop All Motion BIT N RW Ei FM2
'9951 - 9982 User Defined Bits BIT Y RW Eb Ei EN FM1 FM2 MDS
'10001 - 10015 Input Line Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'10017 - 10031 Input Line Raw Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'10033 - 10048 Input Line Debounced Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'10049 - 10064 Output Line Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'10065 - 10096 Input Function Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'10097 - 10128 Output Function Status Array BIT N RO Eb Ei EN FM1 FM2 MDS
'30001 Actual Operating Mode ENM N RO Eb Ei EN FM1 FM2 MDS
'30002 Segment Display Character US16 N RO Eb Ei EN FM1 FM2 MDS
'30004 Actual Operating Mode Expanded ENM N RO Eb EN FM1 MDS
'30101 Input Lines Status Bit Map BM16 N RO Eb Ei EN FM1 FM2 MDS
'30102 Input Lines Raw Status Bit Map BM16 N RO Eb Ei EN FM1 FM2 MDS
'30103 Input Lines Debounced Status Bit Map BM16 N RO Eb Ei EN FM1 FM2 MDS
'30104 Output Lines Status Bit Map BM16 N RO Eb Ei EN FM1 FM2 MDS
'30105 - 30106 Input Function Status Bit Map 32 BM16 N RO Eb Ei EN FM1 FM2 MDS
'30107 - 30108 Output Function Status Bit Map 32 BM16 N RO Eb Ei EN FM1 FM2 MDS
'30401 - 30402 Fault Status Bit Map BM32 N RO Eb Ei EN FM1 FM2 MDS
'31001 Fault 10 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31002 Fault 10 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31003 - 31004 Fault 10 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31005 Fault 9 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31006 Fault 9 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31007 - 31008 Fault 9 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31009 Fault 8 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31010 Fault 8 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31011 - 31012 Fault 8 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31013 Fault 7 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31014 Fault 7 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31015 - 31016 Fault 7 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31017 Fault 6 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31018 Fault 6 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31019 - 31020 Fault 6 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31021 Fault 5 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31022 Fault 5 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31023 - 31024 Fault 5 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31025 Fault 4 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31026 Fault 4 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31027 - 31028 Fault 4 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31029 Fault 3 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31030 Fault 3 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31031 - 31032 Fault 3 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31033 Fault 2 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31034 Fault 2 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31035 - 31036 Fault 2 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'31037 Fault 1 Type ENM Y RO Eb Ei EN FM1 FM2 MDS
'31038 Fault 1 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'31039 - 31040 Fault 1 Power Up Time US32 minutes Y RO Eb Ei EN FM1 FM2 MDS
'32001 - 32002 Pulse Position Input S32 counts N RO Eb EN FM1 MDS
'32021 - 32022 Velocity Feedback S32 RPM N ±13000 0.1 RO Eb Ei EN FM1 FM2 MDS
'32023 Position Feedback (fractional part) US16 revs N 0~0.9999 0.0001 RO Eb EN FM1 MDS
'32024 - 32025 Position Feedback (integral part) S32 revs N RO Eb EN FM1 MDS
'32026 - 32027 Position Feedback S32 revs N ±214748.3647 0.0001 RO Eb Ei EN FM1 FM2 MDS
'32028 - 32029 Following Error S32 revs N ±10 0.0001 RO Eb Ei EN FM1 FM2 MDS
'32032 Shunt Power RMS US16 % N 0~120 0.1 RO Eb Ei EN FM1 FM2 MDS
'32033 Foldback RMS US16 % cont N 0~300 0.1 RO Eb Ei EN FM1 FM2 MDS
'32034 Torque Command S16 % cont N ±300 0.1 RO Eb Ei EN FM1 FM2 MDS
'32035 Torque Command Actual S16 % cont N ±300 0.1 RO Eb Ei EN FM1 FM2 MDS
'32036 - 32037 Position Command S32 revs N ±214748.3647 0.0001 RO Eb Ei EN FM1 FM2 MDS
'32038 Commutation Angle Correction S16 Degrees N ±180 1 RO Eb Ei EN FM1 FM2 MDS
'32039 Commutation Track Angle US16 Degrees N 0~359 1 RO Eb Ei EN FM1 FM2 MDS
'32040 Commutation Voltage S16 % N ±200 0.1 RO Eb Ei EN FM1 FM2 MDS
'32041 Heatsink RMS US16 % N 0~200 0.1 RO Eb Ei EN FM1 FM2 MDS
'32042 Bus Voltage US16 Volts N 20~500 0.1 RO Eb Ei EN FM1 FM2 MDS
'32061 - 32062 Velocity Command S32 RPM N ±13000 0.1 RO Eb Ei EN FM1 FM2 MDS
'32063 Motion State ENM N RO Ei FM2
'32063 - 32064 Velocity Command Analog S32 RPM N ±11000 0.1 RO Eb Ei EN FM1 FM2 MDS
'32065 - 32066 Velocity Command Preset S32 RPM N ±11000 0.1 RO Eb EN FM1 MDS
'32101 Analog Input S16 Volts N ±10 0.001 RO Eb EN FM1 MDS
'32103 Analog Output - Channel 1 S16 Volts N ±10 0.01 RO Eb Ei EN FM1 FM2 MDS
'32104 Analog Output - Channel 2 S16 Volts N ±10 0.01 RO Eb Ei EN FM1 FM2 MDS
'39952 - 39957 FM Firmware Part Number STR N RO FM1 FM2
'39982 Product Group US16 N RO Eb Ei EN FM1 FM2 MDS
'39983 Product Sub-Group US16 N RO Eb Ei EN FM1 FM2 MDS
'39984 Product ID US16 N RO Eb Ei EN FM1 FM2 MDS
'39985 Option 1 ID (Function Module) US16 N RO Eb Ei EN FM1 FM2 MDS
'39988 - 39989 Firmware Revision STR N RO Eb Ei EN MDS
'39990 - 39991 FM Firmware Revision Option STR N RO Eb Ei EN FM1 FM2 MDS
'40001 Operating Mode Default ENM Y RW Eb Ei EN FM1 FM2 MDS
'40002 Motor Type ENM Y RW Eb Ei EN FM1 FM2 MDS
'40003 Axis Address US16 Y 1~99 1 RW Eb Ei EN FM1 FM2 MDS
'40004 Baud Rate ENM Baud Y RW Eb Ei EN FM1 FM2 MDS
'40005 - 40016 Axis Name STR Y RW Eb Ei EN FM1 FM2 MDS
'40018 Operating Mode Default Expanded ENM Y RW Eb EN FM1 MDS
'40019 Operating Mode Alternate ENM Y RW Eb EN FM1 MDS
'40020 Power Up Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40021 - 40022 Power Up Time Total US32 hours Y 0~429496729.5 0.1 RO Eb Ei EN FM1 FM2 MDS
'40023 - 40024 Power Up Time US32 minutes N RO Eb Ei EN FM1 FM2 MDS
'40051 Predefined Setup ENM N RW Eb EN FM1 MDS
'40081 - 40082 Position Feedback Encoder S32 counts N RW Eb Ei EN FM1 FM2 MDS
'40101 Input Lines Force On/Off Command Bit Map US16 N RW Eb Ei EN FM1 FM2 MDS
'40102 Input Lines Force On/Off Enable Bit Map US16 N RW Eb Ei EN FM1 FM2 MDS
'40103 Output Lines Force On/Off Command Bit Map US16 N RW Eb Ei EN FM1 FM2 MDS
'40104 Output Lines Force On/Off Enable Bit Map US16 N RW Eb Ei EN FM1 FM2 MDS
'40105 Output Lines Active Off Bit Map BM16 Y RW Eb Ei EN FM1 FM2 MDS
'40111 - 40123 Input Line Debounce Time US16 Y 0~2000 0.1 RW Eb Ei EN FM1 FM2 MDS
'40201 - 40232 Input Function Mapping 32 ENM Y RW Eb Ei EN FM1 FM2 MDS
'40301 - 40302 Input Function Active Off Bit Map 32 BM16 Y RW Eb Ei EN FM1 FM2 MDS
'40401 - 40402 Input Function Always Active Bit Map 32 BM16 Y RW Eb Ei EN FM1 FM2 MDS
'40451 - 40482 Output Function Mapping 32 US16 Y RW Eb Ei EN FM1 FM2 MDS
'40601 Analog Input Zero Offset S16 Volts Y ±10 0.001 RW Eb EN FM1 MDS
'40602 Analog Input Full Scale S16 Volts Y ±10 0.001 RW Eb EN FM1 MDS
'40603 Analog Input Bandwidth US16 Hz Y 1~1000 1 RW Eb EN FM1 MDS
'40604 Full Scale Velocity US16 RPM Y 0~11000 1 RW Eb EN FM1 MDS
'40605 Full Scale Torque US16 % cont Y 1~300 0.1 RW Eb EN FM1 MDS
'40651 Analog Output 1 Select ENM Y RW Eb Ei EN FM1 FM2 MDS
'40652 - 40653 Analog Output 1 Offset S32 Y ±2147483647 1 RW Eb Ei EN FM1 FM2 MDS
'40654 - 40655 Analog Output 1 Scale S32 Y ±2147483647 1 RW Eb Ei EN FM1 FM2 MDS
'40656 Analog Output 2 Select ENM Y RW Eb Ei EN FM1 FM2 MDS
'40657 - 40658 Analog Output 2 Offset S32 Y ±2147483647 1 RW Eb Ei EN FM1 FM2 MDS
'40659 - 40660 Analog Output 2 Scale S32 Y ±2147483647 1 RW Eb Ei EN FM1 FM2 MDS
'40701 ENcoder State Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40702 ENcoder H/W Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40703 Power Stage Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40704 Low DC Bus Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40705 High DC Bus Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40707 Travel Limit + Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40708 Travel Limit - Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40709 Overspeed Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40711 Power Up Self Test Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40712 NVM Invalid Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40713 Following Error Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40714 Shunt Power RMS Fault Count US16 counts Y RO EN FM1 FM2
'40715 Motor Overtemp Fault Count US16 counts Y RO Eb Ei EN FM1 FM2 MDS
'40716 Drive Overtemp Fault Count US16 counts Y RO Eb Ei MDS
'41001 Pulse Mode Ratio Revolutions S16 revs Y ±2 0.0001 RW Eb EN FM1 MDS
'41002 Pulse Mode Ratio Pulses US16 counts Y 1~16384 1 RW Eb EN FM1 MDS
'41003 Pulse Input Source Select ENM Y RW Eb EN FM1 MDS
'41004 Pulse Interpretation ENM Y RW Eb EN FM1 MDS
'41101 Home Reference ENM Y RW Ei FM2
'41101 - 41102 Velocity Preset 0 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41102 - 41103 Home Velocity S32 RPM Y ±11000 0.1 RW Ei FM2
'41103 - 41104 Velocity Preset 0 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41104 - 41105 Home Acceleration US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'41105 - 41106 Velocity Preset 1 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41106 - 41107 Home Deceleration US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'41107 - 41108 Velocity Preset 1 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41108 - 41109 ENd of Home Position S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'41109 - 41110 Velocity Preset 2 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41110 - 41111 Home Offset S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'41111 - 41112 Velocity Preset 2 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41112 Home Offset Enable ENM Y RW Ei FM2
'41113 - 41114 Home Limit Distance US32 revs Y 0~214748.3647 0.0001 RW Ei FM2
'41113 - 41114 Velocity Preset 3 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41115 Home Limit Distance Enable ENM Y RW Ei FM2
'41115 - 41116 Velocity Preset 3 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41116 Back Off Sensor Before Homing ENM Y RW Ei FM2
'41117 - 41118 Velocity Preset 4 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41119 - 41120 Velocity Preset 4 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41121 - 41122 Velocity Preset 5 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41123 - 41124 Velocity Preset 5 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41125 - 41126 Velocity Preset 6 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41127 - 41128 Velocity Preset 6 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41129 - 41130 Velocity Preset 7 S32 RPM Y ±11000 0.1 RW Eb EN FM1 MDS
'41131 - 41132 Velocity Preset 7 Accel/Decel US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41151 - 41152 Jog Velocity US32 RPM Y 0~11000 0.1 RW Ei FM2
'41153 - 41154 Jog Fast Velocity US32 RPM Y 0~11000 0.1 RW Ei FM2
'41155 - 41156 Jog Acceleration US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'41157 - 41158 Jog Deceleration US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'41201 - 41202 Stop Deceleration US32 ms/kRPM Y 1~32700 0.1 RW Eb Ei EN FM1 FM2 MDS
'41203 - 41204 Travel Limit Deceleration US32 ms/kRPM Y 1~5000 0.1 RW Eb Ei EN FM1 FM2 MDS
'41205 - 41206 Analog Input Accel/Decel Limit US32 ms/kRPM Y 0~32700 0.1 RW Eb EN FM1 MDS
'41306 Motion Command 0 ENM Y RW Ei FM2
'41311 Motion Command 1 ENM Y RW Ei FM2
'41316 Motion Command 2 ENM Y RW Ei FM2
'41321 Motion Command 3 ENM Y RW Ei FM2
'41326 Motion Command 4 ENM Y RW Ei FM2
'41331 Motion Command 5 ENM Y RW Ei FM2
'41336 Motion Command 6 ENM Y RW Ei FM2
'41341 Motion Command 7 ENM Y RW Ei FM2
'41401 Torque Preset 0 S16 % Y ±300 0.1 RW FM1
'41402 Torque Preset 1 S16 % Y ±300 0.1 RW FM1
'41403 Torque Preset 2 S16 % Y ±300 0.1 RW FM1
'41404 Torque Preset 3 S16 % Y ±300 0.1 RW FM1
'41405 Torque Preset 4 S16 % Y ±300 0.1 RW FM1
'41406 Torque Preset 5 S16 % Y ±300 0.1 RW FM1
'41407 Torque Preset 6 S16 % Y ±300 0.1 RW FM1
'41408 Torque Preset 7 S16 % Y ±300 0.1 RW FM1
'42002 Line Voltage ENM Y RW Eb Ei EN FM1 FM2 MDS
'42003 Custom Motor Flag ENM Y RO Eb Ei EN FM1 FM2 MDS
'42021 Inertia Ratio US16 Y 0~50 0.1 RW Eb Ei EN FM1 FM2 MDS
'42023 Friction US16 % cont Y 0~100 0.01 RW Eb Ei EN FM1 FM2
'42024 Response US16 Y 1~500 1 RW Eb Ei EN FM1 FM2 MDS
'42025 High Performance Gains Enable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42026 Feedforwards Enable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42028 Position Error Integral Enable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42029 Position Error Integral Time Constant US16 ms Y 5~500 1 RW Eb Ei EN FM1 FM2 MDS
'42031 Following Error Enable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42032 - 42033 Following Error Limit S32 revs Y 0.001~10 0.0001 RW Eb Ei EN FM1 FM2 MDS
'42034 Torque Limit US16 % cont Y 0~300 1 RW Eb Ei EN FM1 FM2 MDS
'42035 In Motion Velocity US16 RPM Y 0~100 1 RW Eb Ei EN FM1 FM2 MDS
'42036 Overspeed Velocity Limit US16 RPM Y 0~13000 1 RW Eb Ei EN FM1 FM2 MDS
'42044 Positive Direction ENM Y RW Eb Ei EN FM1 FM2 MDS
'42045 Drive Ambient Temperature US16 Degree C Y 20~50 1 RW Eb Ei EN FM1 FM2 MDS
'42046 Low DC Bus Enable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42047 Low Pass Filter Enable (COMPFE) ENM Y RW Eb Ei EN FM1 FM2 MDS
'42048 Low Pass Filter Frequency (COMPF) US16 Hz Y 1-1000 1 RW Eb Ei EN FM1 FM2 MDS
'42049 Torque Level 1 (MSTL1) US16 %cont Y 0-300 1 RW Eb Ei EN FM1 FM2 MDS
'42050 Torque Level 2 (MSTL2) US16 %cont Y 0-300 1 RW Eb Ei EN FM1 FM2 MDS
'42061 ENcoder Output Scaling US16 lines/rev Y 1~8192 1 RW Eb Ei EN FM1 FM2 MDS
'42062 ENcoder Output Scaling ENable ENM Y RW Eb Ei EN FM1 FM2 MDS
'42101 - 42106 User Defined Motor Name STR Y RO Eb Ei EN FM1 FM2 MDS
'42107 Motor Poles ENM Y RO Eb Ei EN FM1 FM2 MDS
'42108 Motor Encoder Lines Per Revolution ENM Lines Y RO Eb Ei EN FM1 FM2 MDS
'42109 Motor Encoder Marker Angle US16 Degrees Y 0~359 1 RO Eb Ei EN FM1 FM2 MDS
'42110 Motor Encoder U Angle US16 Degrees Y 0~359 1 RO Eb Ei EN FM1 FM2 MDS
'42111 Motor Encoder Reference Motion ENM Y RO Eb Ei EN FM1 FM2 MDS
'42112 Motor Inertia US16 Y 0.00001~0.5 0.00001 RO Eb Ei EN FM1 FM2 MDS
'42113 Motor KE US16 vrms/kRPM Y 5~500 0.1 RO Eb Ei EN FM1 FM2 MDS
'42114 Motor Resistance US16 Ohms Y 0.1~50 0.01 RO Eb Ei EN FM1 FM2 MDS
'42115 Motor Inductance US16 mH Y 1~100 0.1 RO Eb Ei EN FM1 FM2 MDS
'42116 Motor Continuous Current Rating US16 Arms Y 0.1~100 0.01 RO Eb Ei EN FM1 FM2 MDS
'42117 Motor Peak Current Rating US16 Arms Y 1~100 0.01 RO Eb Ei EN FM1 FM2 MDS
'42118 Motor Maximum Operating Speed US16 RPM Y 0~11000 1 RO Eb Ei EN FM1 FM2 MDS
'43001 Index Type 0 ENM Y RW Ei FM2
'43002 - 43003 Index Distance 0 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43004 - 43005 Index Velocity 0 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43006 - 43007 Index Acceleration 0 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43008 - 43009 Index Deceleration 0 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43026 Index Type 1 ENM Y RW Ei FM2
'43027 - 43028 Index Distance 1 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43029 - 43030 Index Velocity 1 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43031 - 43032 Index Acceleration 1 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43033 - 43034 Index Deceleration 1 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43051 Index Type 2 ENM Y RW Ei FM2
'43052 - 43053 Index Distance 2 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43054 - 43055 Index Velocity 2 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43056 - 43057 Index Acceleration 2 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43058 - 43059 Index Deceleration 2 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43076 Index Type 3 ENM Y RW Ei FM2
'43077 - 43078 Index Distance 3 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43079 - 43080 Index Velocity 3 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43081 - 43082 Index Acceleration 3 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43083 - 43084 Index Deceleration 3 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43101 Index Type 4 ENM Y RW Ei FM2
'43102 - 43103 Index Distance 4 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43104 - 43105 Index Velocity 4 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43106 - 43107 Index Acceleration 4 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43108 - 43109 Index Deceleration 4 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43126 Index Type 5 ENM Y RW Ei FM2
'43127 - 43128 Index Distance 5 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43129 - 43130 Index Velocity 5 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43131 - 43132 Index Acceleration 5 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43133 - 43134 Index Deceleration 5 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43151 Index Type 6 ENM Y RW Ei FM2
'43152 - 43153 Index Distance 6 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43154 - 43155 Index Velocity 6 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43156 - 43157 Index Acceleration 6 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43158 - 43159 Index Deceleration 6 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43176 Index Type 7 ENM Y RW Ei FM2
'43177 - 43178 Index Distance 7 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43179 - 43180 Index Velocity 7 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43181 - 43182 Index Acceleration 7 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43183 - 43184 Index Deceleration 7 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43201 Index Type 8 ENM Y RW Ei FM2
'43202 - 43203 Index Distance 8 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43204 - 43205 Index Velocity 8 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43206 - 43207 Index Acceleration 8 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43208 - 43209 Index Deceleration 8 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43226 Index Type 9 ENM Y RW Ei FM2
'43227 - 43228 Index Distance 9 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43229 - 43230 Index Velocity 9 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43231 - 43232 Index Acceleration 9 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43233 - 43234 Index Deceleration 9 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43251 Index Type 10 ENM Y RW Ei FM2
'43252 - 43253 Index Distance 10 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43254 - 43255 Index Velocity 10 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43256 - 43257 Index Acceleration 10 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43258 - 43259 Index Deceleration 10 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43276 Index Type 11 ENM Y RW Ei FM2
'43277 - 43278 Index Distance 11 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43279 - 43280 Index Velocity 11 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43281 - 43282 Index Acceleration 11 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43283 - 43284 Index Deceleration 11 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43301 Index Type 12 ENM Y RW Ei FM2
'43302 - 43303 Index Distance 12 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43304 - 43305 Index Velocity 12 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43306 - 43307 Index Acceleration 12 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43308 - 43309 Index Deceleration 12 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43326 Index Type 13 ENM Y RW Ei FM2
'43327 - 43328 Index Distance 13 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43329 - 43330 Index Velocity 13 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43331 - 43332 Index Acceleration 13 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43333 - 43334 Index Deceleration 13 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43351 Index Type 14 ENM Y RW Ei FM2
'43352 - 43353 Index Distance 14 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43354 - 43355 Index Velocity 14 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43356 - 43357 Index Acceleration 14 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43358 - 43359 Index Deceleration 14 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43376 Index Type 15 ENM Y RW Ei FM2
'43377 - 43378 Index Distance 15 S32 revs Y ±214748.3647 0.0001 RW Ei FM2
'43379 - 43380 Index Velocity 15 US32 RPM Y 0~11000 0.1 RW Ei FM2
'43381 - 43382 Index Acceleration 15 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'43383 - 43384 Index Deceleration 15 US32 ms/kRPM Y 1~32700 0.1 RW Ei FM2
'49401 - 49402 User Defined Bitmap US16 Y RW Eb Ei EN FM1 FM2 MDS
'49403 - 49418 User Defined Registers US16 Y RW Eb Ei EN FM1 FM2 MDS
'49903 - 49910 Product Serial Number STR Y RO Eb Ei EN FM1 FM2 MDS
'49957 - 49962 FM Serial Number STR N RO FM1 FM2
'
Public Declare Function NTBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long
Global eTitle$
Global eMess$
Global mError As Long
Global Ser$
Global PortNo
Global QUIT
Global Maxtry
Global Retry
Global Repeat
Global RX$
Global RXRaw$
Global OLDRX$
Global PGCHK
Global DEVTO
Global TRIES
Global T$
Global Send$
Global NodeAddr As Integer
Global Addr$
Global AddrIn As Integer
Global Qty$
Global LoQty$
Global HiQty$
Global ExpQty As Integer
Global Func$
Global Slave$
Global HiAddr$
Global LoAddr$
'
Global FuncIn$
Global SlaveIn$
Global HiAddrIn$
Global LoAddrIn$
Global QtyIn As Byte
Global LoQtyIn$
Global HiQtyIn$
Global LoCRCIn$
Global HiCRCIn$
Global DataByteIn$
Global CRCIn As Long
'
Global InStat$
Global In1$
Global In2$
Global In3$
Global In1Lo$
Global In1Hi$
Global In2Lo$
Global In2Hi$
Global In3Lo$
Global In3Hi$
Global CHK$
Global DECVal
Global WordIn As Integer
Global WordInStr$
Global bitMax As Integer
Global Bitw As Integer
Global BitNo As Integer
Global Bit As Integer
Global MaxReg As Integer
'
Global ExErr$(6)
Global FL(12)
Global Coil(250)
Global Inputs(64)
Global Outputs(64)
Global ModIn(1000) 'DIM SHARED ARRAY FOR Inputs 0-47
Global ByteTen As Integer
Global ComErr As Byte
Global Com1Err As Byte
Global i As Integer
Global j As Integer
Global Reg$
Global Reg2$
Global RxIn() As Integer 'dimension Array For All Data
Global Target$
Global HRBase As Long
Dim Ans$
Dim SENDTM$
Dim OLDTM!
Dim NEWTM!
Dim TMTRY!
Dim AddErr As Integer
Dim RxChar$
Dim CharInS As Integer
Dim CommIn$
Dim CommInLen As Integer
Dim RXHex As Integer
Dim RxLen As Integer
Dim BufLen As Integer
Dim Buffer$
Dim ChkErr$
Dim AddError$
Dim RxMRK$
Dim ExCHK$
Dim ExCODE$
Dim CODE
Dim ChkSum As Integer
Global Fails As Integer
Global CommDone As Byte
Global CommComp As Byte
Global Result
Global OffLine As Byte
Global RxInTOT$
Global RxTOT$
Global msg$
Global ERRCHAR$
Global Reply$
Global BCCSum
Global BCCSumStr$
Global CRC
Global CRCRegLo As Byte
Global CRCRegHi As Long
Global CRCReg As Long
Global CRCStr$
Global CRCFrame$
Global AppDir
Global AppName$
Global Mess$
Global Continuous As Byte
Global OutData As Integer
Global OutReg As Integer
Global nobits As Byte
Dim Confirm As Long
'below for evan toder's trick to make a command button stay "in" when pressed
Global Const BM_SETSTATE = &HF3
Public Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Sub Main()
XPMain
Load frmMain
frmMain.Show
End Sub

Public Sub DecDat()
On Error GoTo Oops
Dim k As Integer
'below is from slc
'***************************************************************
'SUBROUTINE TO ISOLATE INPUT DATA
WordInStr$ = Right$((Hex$(WordIn)), 4)
WordInStr$ = AddZero(WordInStr$, 4)
InStat$ = Right$(WordInStr$, 4)
For k = 0 To 15
    Bit = BitNo + k
    ModIn(Bit) = ((2 ^ k) And WordIn) / (2 ^ k)
Next k
GoTo Exit_DecDat
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine DecDat "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in DecDat"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_DecDat:
End Sub

Sub ReadCoils(BegAddress As String, RegQty As String)
'READ COIL STATUS
'*************************************************************************
Func$ = "01"
'THIS FUNCTION WILL READ THE STATUS OF 32 CONSECUTIVE COILS."
Addr$ = BegAddress ' 0XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 9999) Then
    Beep
    MsgBox Addr$ & " IS OUT OF RANGE"
    Exit Sub
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " IS INVALID."
    Exit Sub
End If
HRBase = Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
Qty$ = RegQty
bitMax = Val(RegQty)
ExpQty = (bitMax \ 8) + (IIf((bitMax Mod 8) = 0, 0, 1))
Target$ = "COIL "
'SEND MESSAGE
Call TalkRTU
Call ShowIn
End Sub

Sub ReadInputs(BegAddress As String, RegQty As String)
'READ DISCRETE INPUTS.  Up to 2048 inputs may be read at one pass.
'*********************************************************************
Func$ = "02"
'THIS FUNCTION WILL READ THE STATUS OF 48 CONSECUTIVE INPUTS."
Addr$ = BegAddress ' 0XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 9999) Then
    Beep
    MsgBox Addr$ & " Is Out Of Range"
    Exit Sub
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " Is Invalid."
    Exit Sub
End If
HRBase = 10000 + Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
Target$ = "INPUT "
Qty$ = RegQty
bitMax = Val(RegQty)
'we need to round up to the next byte
ExpQty = (bitMax \ 8) + (IIf((bitMax Mod 8) = 0, 0, 1))
'SEND MESSAGE AND DISPLAY THE REPLY (Message is created here).
Call TalkRTU
Target$ = "INPUT "
'WAIT FOR REPLY AND THEN DISPLAY
If ComErr = 0 Then Call ShowIn
End Sub

Sub ReadReg(BegAddress As String, RegQty As String)
'READ HOLDING REGISTERS.  up to 125 holding registers may be read.
REGSTAT:
'THIS FUNCTION WILL READ THE CONTENTS OF UP TO 125 HOLDING REGISTERS."
Func$ = "03"
REGIN:
Addr$ = BegAddress ' 4XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 49999!) Then
    Beep
    MsgBox Addr$ & " IS OUT OF RANGE"
    Exit Sub
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " IS INVALID."
    Exit Sub
End If
If Len(Addr$) > 4 Then Addr$ = Right$(Addr$, 4)
HRBase = 40000 + Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
Qty$ = RegQty
ExpQty = 2 * Val(RegQty)
If Val(Qty$) > 125 Then
    Beep
    MsgBox " No more than 125 registers may be read at a time."
    Exit Sub
ElseIf (Qty$ = "") Or (Val(Qty$) = 0) Then
    Exit Sub
End If
'THIS FUNCTION WILL READ THE CONTENTS OF UP TO 125 HOLDING REGISTERS."
Call TalkRTU
RegDec
End Sub

Sub ReadInReg(BegAddress As String, RegQty As String)
'******************************************************************************************************************
'READ INPUT REGISTERS.
' "       THIS FUNCTION WILL READ THE CONTENTS OF THE INPUT REGISTERS."
' The first character has the high bit and the second has the low.
Func$ = "04"
Addr$ = BegAddress ' 0XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 9999) Then
    Beep
    Addr$ = Right$(Addr$, 4)
    If (Len(Addr$) > 5) Or (Val(Addr$) > 9999) Then
        Beep
        MsgBox Addr$ & " IS OUT OF RANGE"
        Exit Sub
    End If
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " IS INVALID."
    Exit Sub
End If
HRBase = 30000 + Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
Qty$ = RegQty
ExpQty = 2 * Val(RegQty)
Call TalkRTU
RegDec
End Sub

Sub SetOutput(CoilAddress As String, CoilStat As Byte)
' MODIFY AN OUTPUT COIL.  The coil may be turned either ON or OFF
'THIS FUNCTION WILL MODIFY THE STATUS OF ONE COIL."
Func$ = "05"  'The function code for modifying an output coil condition
Addr$ = CoilAddress '0XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 9999) Then
    Beep
    MsgBox Addr$ & " IS OUT OF RANGE"
    Exit Sub
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " IS INVALID."
    Exit Sub
End If
If Len(Addr$) > 4 Then Addr$ = Right$(Addr$, 4)
HRBase = Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
'PRINT "RETRIEVING STATUS OF COIL "; ADDR;
Func$ = "01"
Qty$ = "1"
ExpQty = 1
Call TalkRTU
'Enter (1) to turn the coil ON, or (0) to turn it OFF. "
If CoilStat = 0 Then
    Qty$ = "0000"
    LoQty$ = "00"
    HiQty$ = "00"
ElseIf CoilStat = 1 Then
    Qty$ = "FF00"
    LoQty$ = "00"
    HiQty$ = "FF"
Else
    Beep
    Exit Sub
End If
Func$ = "05"
Call TalkRTU
'WAIT FOR REPLY AND THEN DISPLAY
RegDec
End Sub

Sub WriteReg(BegAddress As String, NewVal As String)
'THIS FUNCTION WILL MODIFY THE CONTENTS OF ONE HOLDING REGISTER."
Func$ = "06"  'The function code for modifying a holding register.
Addr$ = BegAddress ' 4XXXX
If (Len(Addr$) > 5) Or (Val(Addr$) > 49999!) Then
    Beep
    MsgBox Addr$ & " IS OUT OF RANGE"
    Exit Sub
ElseIf (Addr$ = "") Or (Val(Addr$) < 1) Then
    Beep
    MsgBox Addr$ & " IS INVALID."
    Exit Sub
End If
If Len(Addr$) > 4 Then Addr$ = Right$(Addr$, 4)
HRBase = 40000 + Val(Addr$) - 1
Addr$ = AddZero(Hex$(Val(Addr$) - 1), 4)
Qty$ = NewVal '" Enter the new value to hold, ( decimal, 0 - 32767 ): ")
ExpQty = 2
If Qty$ = "" Then
    Exit Sub
End If
Qty$ = Str$(Val(Qty$))
Call TalkRTU
'WAIT FOR REPLY AND THEN DISPLAY
End Sub

Sub ReadException()
'Reads the contents of eight Exception Status coils within the slave controller.
'Certain coils have predefined assignments in the various controllers. Other coils
'can be programmed by the user to hold information about the contoller’s status,
'for example, ‘machine ON/OFF’, ‘heads retracted’, ‘safeties satisfied’, ‘error
'conditions exist’, or other user–defined flags. Broadcast is not supported.
'The function provides a simple method for accessing this information, because the
'Exception Coil references are known (no coil reference is needed in the function).
'The predefined Exception Coil assignments are:
'Controller Model Coil Assignment
'M84, 184/384, 584, 984 1 – 8 User defined
'484 257 Battery Status
'258 – 264 User defined
'884 761 Battery Status
'762 Memory Protect Status
'763 RIO Health Status
'764–768 User defined
'
'The normal response contains the status of the eight Exception Status coils.
'The coils are packed into one data byte, with one bit per coil. The status of the
'lowest coil reference is contained in the least significant bit of the byte.
'
'This Function Reads The Contents Of Eight Exception Status Coils
Func$ = "07"
Addr$ = ""
Qty$ = ""
Call TalkRTU
End Sub

Sub RegDec()
'************************************************************************
' THIS IS AN EXAMPLE PORTION OF CODE TO DISPLAY THE DECIMAL CONTENTS
'OF HOLDING OR INPUT REGISTERS
Dim RegNo As Integer
MaxReg = AddrIn + (ExpQty \ 2) + 1
Erase RxIn()
ReDim RxIn(MaxReg)
If ComErr = 1 Then Exit Sub
'the byte quantity is the 10th byte
'calculates the number of data bytes received
ByteTen = Val("&H" + Mid$(RX$, 5, 2))
For i = 7 To ((ByteTen * 2) + 3) Step 4
    'data starts at byte 7 of the returned message
    Reg$ = Mid$(RX$, i, 2)
    Reg2$ = Mid$(RX$, i + 2, 2)
    'calculate the decimal value of the register
    'holding registers have the high byte first
    'bit data has the low byte first
    If Val(Func$) = 1 Or Val(Func$) = 2 Then
        'swap the bytes for coil and input data
        DECVal = Val("&H" & Reg2$ & Reg$)
    Else
        DECVal = Val("&H" & Reg$ & Reg2$)
    End If
    RegNo = AddrIn + ((i - 3) / 4)
    RxIn(RegNo) = DECVal
    'Debug.Print "The Decimal value of register "; (HRBase + (i - 3) / 4); " is "; DECVal; " ("; Reg$; Reg2$; "H)"
    'Debug.Print i,
Next i
'Debug.Print
'Debug.Print i
If RegNo < MaxReg And ExpQty Mod 2 = 1 Then
    'we have an odd number of bytes
    'this happens if we are reading coils or inputs
    'Debug.Print Len(RX$)
    Reg$ = Mid$(RX$, i, 2)
    DECVal = Val("&H" & Reg$)
    RxIn(MaxReg) = DECVal
End If
End Sub

Public Sub ShowIn()
On Error GoTo Oops
RegDec
'SUBROUTINE FOR DECIPHERING DATA
For Bitw = 0 To ((ExpQty / 2) - 1)
    'pick out the registers
    WordIn = (RxIn(AddrIn + Bitw + 1))
    'isolate the bits starting at bitno
    BitNo = (AddrIn * 16) + (Bitw * 16) + 1
    Call DecDat
Next Bitw
'uncomment to debug the bit status
'PRINT INPUTS
'For i = 0 To ExpQty - 1
'    For j = 1 To 16
'        BitNo = (AddrIn * 16) + (i * 16) + j
'        Debug.Print BitNo; "="; ModIn(BitNo - 1); " - ";
'    Next j
'    Debug.Print
'Next i
GoTo Exit_ShowIn
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ShowIn "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ShowIn"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ShowIn:
End Sub

Public Sub TalkRTU()
On Error GoTo Oops
Dim DLTries As Integer
Dim HexSend$
'Command Description
'1 Read Coil Status
'2 Read Input Status
'3 Read Holding Registers
'4 Read Inputs Registers
'5 Force Single Coil
'6 Preset Single Register
'7 Read Exception Status
'8 Perform Diagnostic Test
'15 Force Multiple Coils
'16 Preset Multiple Registers
'17 Report Slave ID number
'
Screen.MousePointer = vbHourglass
CommDone = 0
CommComp = 0
Maxtry = 5
RX$ = ""
'*************************************************************************
'CREATE THE Qty$ PART OF THE MESSAGE.
Qty$ = AddZero(Hex$(Val(Qty$)), 4)
If Len(Qty$) > 4 Then Qty$ = Right$(Qty$, 4)
LoQty$ = Right$(Qty$, 2)
HiQty$ = Left$(Qty$, 2)
'*************************************************************************
'
TRIES = 0
' Message is created here.
BuildMess:
'Input the desired slave address.
Slave$ = AddZero(Hex$(NodeAddr), 2)
'Split the Address
LoAddr$ = Right$(Addr$, 2)
HiAddr$ = Left$(Addr$, 2)
'
If Val(Func$) = 7 Then
    Mess$ = Chr$(Val("&H" & Slave$)) & Chr$(Val("&H" & Func$))
    CRCcheck Mess$
    Mess$ = Mess$ & Chr$(CRCRegLo) & Chr$(CRCRegHi)
    GoTo SendMess
End If
Mess$ = Chr$(Val("&H" & Slave$)) & Chr$(Val("&H" & Func$)) & Chr$(Val("&H" & HiAddr$)) & Chr$(Val("&H" & LoAddr$)) & Chr$(Val("&H" & HiQty$)) & Chr$(Val("&H" & LoQty$))
CRCcheck Mess$
'The CRC field is appended to the message as the last field in the message.
'When this is done, the low–order byte of the field is appended first, followed by the
'high–order byte. The CRC high–order byte is the last byte to be sent in the message.
Mess$ = Mess$ & Chr$(CRCRegLo) & Chr$(CRCRegHi)
HexSend$ = Slave$ & Func$ & HiAddr$ & LoAddr$ & HiQty$ & LoQty$ & Hex$(CRCRegLo) & Hex$(CRCRegHi)
'01 03 0000 0001 FB   01      03    0000    0001    FB
ComErr = 0
Com1Err = 0
Send$ = Mess$
Repeat = 1
SendMess:
SENDTM$ = Time$
'-------------------------------------------
TRANSMIT:
If OffLine = 1 Then
    CommDone = 1
    CommComp = 1
    frmMain.Caption = "Tom Pydeski's Modbus Communications - OffLine"
    GoTo endT
    Exit Sub
Else
    frmMain.Caption = "Tom Pydeski's Modbus Communications - OnLine"
End If
OLDTM! = Timer
If frmMain.Comm1.InBufferCount > 0 Then EmptyBuffer
frmMain.Comm1.Output = Send$
tp:
RxInTOT$ = ""
'******************************************************************************************************************
'this Routine Handles The Replies To A Query.
'this Includes Any Necessary Retries Due To Invalid
'checksums Or No Reply After Maxtry Attempts
'
Receeve:
AddErr = 0
' If something In INPUT buffer, reinitialize the received message string.
AddErr = 0
RxTOT$ = ""
CharInS = frmMain.Comm1.InBufferCount
If CharInS = 0 Then
    GoTo TOCHK ' If INPUT buffer is empty go and check if retry time has passed yet.
End If
frmMain.Comm1.InputLen = CharInS
CommIn$ = frmMain.Comm1.Input
CommIn$ = StrConv(CommIn$, vbUnicode)
CommInLen = Len(CommIn$)
'Debug.Print CommIn$
If CommIn$ = Send$ Then
    'we were echoed back what we sent
    'if we modified a register or coil, this is expected
    If Val(Func$) = 5 Or Val(Func$) = 6 Then
        Do While Len(CommIn$) > 0
            Pick
            RX$ = RX$ & AddZero(Hex$(Asc(RxChar$)), 2)
            RXRaw$ = RXRaw$ & RxChar$
        Loop
        msg = msg & vbCrLf & " The Message Sent Was = > " & HexSend$
        msg = msg & vbCrLf & "The Reply Received Was => " & RX$
        EmptyBuffer
        GoTo Exit_Talk
    End If
    MsgBox "We received the same string we sent!"
End If
'Debug.Print CommInLen; " characters in buffer"
'we now have to pick off the characters in the reply
RX$ = ""
RXRaw$ = ""
ColonIN:
DLTries = 0
ColonIN2:
Pick
SlaveIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
RXRaw$ = RXRaw$ & RxChar$
RX$ = RX$ & SlaveIn$
If SlaveIn$ <> Slave$ Then
    DLTries = DLTries + 1
    If DLTries < CharInS Then
        GoTo ColonIN2
    End If
    msg = "Slave$ = " & SlaveIn$ & " => " & Slave$
    GoTo CommError
End If
getfunc:
Pick
FuncIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
RXRaw$ = RXRaw$ & RxChar$
RX$ = RX$ & FuncIn$
If FuncIn$ <> Func$ Then
    DLTries = DLTries + 1
    If DLTries < CharInS Then GoTo getfunc
    msg = "Func$ = " & FuncIn$ & " => " & Func$
    GoTo CommError
End If
getqty:
Pick
LoQtyIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
RXRaw$ = RXRaw$ & RxChar$
RX$ = RX$ & LoQtyIn$
QtyIn = Val("&H" & LoQtyIn$)
If QtyIn <> ExpQty Then
    'we don't have the data we were expecting
    DLTries = DLTries + 1
    msg = "LoQtyIn$ = " & LoQtyIn$ & " => " & ExpQty
    GoTo CommError
End If
'we should now be at the data
If Com1Err = 1 Then TRIES = TRIES + 1: GoTo BuildMess
'--------------------------
GetRest:
For i = 1 To QtyIn
    Pick
    DataByteIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
    'Debug.Print DataByteIn$
    RXRaw$ = RXRaw$ & RxChar$
    RX$ = RX$ & DataByteIn$
Next i
CRCcheck RXRaw$
'CHECKING THE CHECKSUM OF THE REPLY.
'we now have to do the crc
getcrc:
Pick
LoCRCIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
RXRaw$ = RXRaw$ & RxChar$
RX$ = RX$ & LoCRCIn$
Pick
HiCRCIn$ = AddZero(Hex$(Asc(RxChar$)), 2)
RXRaw$ = RXRaw$ & RxChar$
RX$ = RX$ & HiCRCIn$
CRCIn = Val(Abs("&H" & HiCRCIn$ & LoCRCIn$))
If CRCIn <> CRCReg Then
    msg = "CRC Error = " & CRCIn & " - " & CRCReg
End If
If CRCIn = 0 Then GoTo ChkError
'--------------------------
EndRx:
'
'CLEAR INPUT BUFFER
CLEARBUF:
BufLen = frmMain.Comm1.InBufferCount
If BufLen = 0 Then GoTo GotAll
Buffer$ = frmMain.Comm1.Input
If Val(Buffer$) > 0 Then Debug.Print Buffer$; "   "
'we must have an error of some kind
msg$ = "Possible transmission error."
msg$ = msg & vbCrLf & "Now retransmitting original message."
msg$ = msg & vbCrLf & "    The Message Sent Was " & Send$
msg$ = msg & vbCrLf & "The Message Received Was " & RX$
msg$ = msg & vbCrLf & "Checksum Value Was " & AddError$
msg$ = msg & vbCrLf & "The Hex Value Was " & Val("&h" + AddError$)
MsgBox msg, vbCritical, "Communication Error with PLC"
Retry = 1
Com1Err = 1
ChkError:
If Retry Then Retry = 0: GoTo TOCHK
'
GotAll:
'CHECKING FOR AN EXCEPTION ERROR.
RxMRK$ = Mid$(RX$, 1, 4)
ExCHK$ = Mid$(RxMRK$, 3, 1)
If ExCHK$ = "8" Then
    GoTo ExcErr
Else
    DEVTO = 0
    GoTo Exit_Talk
End If
'************************************************************************
ExcErr:
ExCODE$ = Mid$(RX$, 5, 2)
CODE = Val(ExCODE$)
msg$ = " Exception Error #" & CODE & " - " & ExErr$(CODE)
msg$ = msg & vbCrLf & "Sent =>" & Send$
msg$ = msg & vbCrLf & "Received =>" & RX$
msg$ = msg & vbCrLf & " Exception Error #" & CODE & " - " & ExErr$(CODE)
MsgBox msg, vbCritical, "Communication Error with PLC"
GoTo Exit_Talk
'************************************************************************
TOCHK:
'CHECKING FOR EXCESSIVE REPLY TIME.
'If so, a retransmission of the message not responded to will occur.
NEWTM! = Timer:
If (NEWTM! - OLDTM!) < 0 Then NEWTM! = NEWTM! + 86400!
If (NEWTM! - OLDTM!) > TMTRY! Then
    Repeat = Repeat + 1
    Fails = Fails + 1
    If Repeat = Maxtry + 1 Then
        GoTo NoReply
    Else
        GoTo SendMess
    End If
End If
GoTo Receeve
'************************************************************************
NoReply:
ComErr = 1
Repeat = 1
Confirm = MsgBox("No Reply After " & Maxtry & " Attempts" & vbNewLine & "Do you want to run Offline?", vbCritical + vbYesNo + vbMsgBoxSetForeground, "Communication Error with PLC")
Repeat = 1
If Confirm = vbYes Then OffLine = 1
ByteTen = 0
GoTo Exit_Talk
'************************************************************************
CommError:
Do While Len(CommIn$) > 0
    Pick
    RX$ = RX$ & AddZero(Hex$(Asc(RxChar$)), 2)
    RXRaw$ = RXRaw$ & RxChar$
Loop
msg = msg & vbCrLf & " The Message Sent Was = > " & HexSend$
msg = msg & vbCrLf & "The Reply Received Was => " & RX$
MsgBox msg, vbCritical, "Communication Error with PLC"
EmptyBuffer
GoTo Exit_Talk
'************************************************************************
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Talk "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Talk"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Talk:
'REPLY$ = DLEStr$ + ACKStr$
'frmmain.Comm1.Output = REPLY$
GoTo endT
endT:
msg$ = " The Message Sent Was = > " & HexSend$
msg$ = msg$ & vbCrLf & "The Reply Received Was => " & RX$
'Debug.Print msg$
If Continuous = 0 Then
    If frmMain.Text1.Text <> msg$ & vbCrLf Then
        frmMain.Text1.Text = frmMain.Text1.Text & msg$ & vbCrLf
    End If
Else
    If frmMain.Text1.Text <> msg$ & vbCrLf Then
        frmMain.Text1.Text = msg$ & vbCrLf
    End If
End If
CommComp = 1
Screen.MousePointer = 0
End Sub

Public Sub TalkASCII()
On Error GoTo Oops
Dim DLTries As Integer
'Command Description
'1 Read Coil Status
'2 Read Input Status
'3 Read Holding Registers
'4 Read Inputs Registers
'5 Force Single Coil
'6 Preset Single Register
'7 Read Exception Status
'8 Perform Diagnostic Test
'15 Force Multiple Coils
'16 Preset Multiple Registers
'17 Report Slave ID number
'
Screen.MousePointer = vbHourglass
CommDone = 0
CommComp = 0
Maxtry = 5
AddrIn = Val("&H" & Addr$)
'*************************************************************************
'CREATE THE Qty$ PART OF THE MESSAGE.
If Len(Qty$) > 4 Then Qty$ = Right$(Qty$, 4)
Qty$ = AddZero(Hex$(Val(Qty$)), 4)
LoQty$ = Right$(Qty$, 2)
HiQty$ = Left$(Qty$, 2)
'*************************************************************************
'
TRIES = 0
' Message is created here.
BuildMess:
'Input the desired slave address.
Slave$ = AddZero(Hex$(NodeAddr), 2)
'Split the Address
LoAddr$ = Right$(Addr$, 2)
HiAddr$ = Left$(Addr$, 2)
'
'Transmit as ascii
'ERROR CHECKING SUBROUTINE FOR ASCII TRANSMISSION.
ChkSum = Val("&H" & Func$) + Val("&H" & Slave$) + Val("&H" & HiAddr$) + Val("&H" & LoAddr$) + Val("&H" & LoQty$) + Val("&H" & HiQty$)
T$ = Hex$(ChkSum)
'now get 2's complement???
ChkSum = Not ChkSum
CHK$ = Hex$(ChkSum + 1)
CHK$ = Right$(CHK$, 2)
'Debug.Print Val("&H" & CHK$)
Mess$ = Chr$(58) & Slave$ & Func$ & Addr$ & Qty$ & CHK$ & vbCrLf 'Chr$(13)
'01 03 0000 0001 FB   01      03    0000    0001    FB
ComErr = 0
Com1Err = 0
Send$ = Mess$
Repeat = 1
SendMess:
SENDTM$ = Time$
'-------------------------------------------
TRANSMIT:
If OffLine = 1 Then
    CommDone = 1
    CommComp = 1
    GoTo endT
    Exit Sub
End If
OLDTM! = Timer
frmMain.Comm1.Output = Send$
tp:
RxInTOT$ = ""
'******************************************************************************************************************
'this Routine Handles The Replies To A Query.
'this Includes Any Necessary Retries Due To Invalid
'checksums Or No Reply After Maxtry Attempts
'
Receeve:
AddErr = 0
RX$ = ""
' If something In INPUT buffer, reinitialize the received message string.
AddErr = 0
RxTOT$ = ""
CharInS = frmMain.Comm1.InBufferCount
If CharInS = 0 Then
    GoTo TOCHK ' If INPUT buffer is empty go and check if retry time has passed yet.
End If
frmMain.Comm1.InputLen = CharInS
CommIn$ = frmMain.Comm1.Input
CommIn$ = StrConv(CommIn$, vbUnicode)
CommInLen = Len(CommIn$)
RX$ = ""
'Pick off the colon in the message.
ColonIN:
DLTries = 0
Pick
If RxChar$ <> ":" Then
    DLTries = DLTries + 1
    If DLTries < CharInS Then GoTo ColonIN
    msg = "colon = " & RxChar$
    GoTo CommError
End If
Pick
4390  If Asc(RxChar$) = 13 Then GoTo Receeve
RXHex = Val("&H" + RxChar$)
If RXHex = 203 Then GoTo CLEARBUF
If RXHex > 15 Then
    Com1Err = 1
    TRIES = TRIES + 1
    GoTo BuildMess
End If
RX$ = RX$ + RxChar$
EndRx:
'
'debug.PRINT RX$; "                   "
'CLEAR INPUT BUFFER
CLEARBUF:
BufLen = frmMain.Comm1.InBufferCount
If BufLen = 0 Then GoTo 4475
Buffer$ = frmMain.Comm1.Input
If Val(Buffer$) > 0 Then Debug.Print Buffer$; "   "
4475 'CHECKING THE CHECKSUM OF THE REPLY.
4480     BufLen = Len(RX$)
SKIPT:  If BufLen < 3 Then GoTo RetrySend
4485     For i = 1 To BufLen - 1 Step 2
4487        ChkErr$ = Mid$(RX$, i, 2)
AddErr = AddErr + Val("&H" + ChkErr$)
4490     Next i
4499     AddError$ = Right$(Hex$(AddErr), 2)
4500     If Val("&H" + AddError$) = 0 Then GoTo 4510
'we must have an error of some kind
msg$ = "Possible transmission error."
msg$ = msg & vbCrLf & " Now Retransmitting Original Message."
msg$ = msg & vbCrLf & " The Message Sent Was " & Send$
msg$ = msg & vbCrLf & " And The Message Received Was " & RX$
msg$ = msg & vbCrLf & " Checksum Value Was " & AddError$
msg$ = msg & vbCrLf & " The Hex Value Was " & Val("&h" + AddError$)
MsgBox msg, vbCritical, "Communication Error with PLC"
RetrySend:
Retry = 1
Com1Err = 1
4510     If Retry Then Retry = 0: GoTo TOCHK
'CHECKING FOR AN EXCEPTION ERROR.
RxMRK$ = Mid$(RX$, 1, 4)
ExCHK$ = Mid$(RxMRK$, 3, 1)
If ExCHK$ = "8" Then
    GoTo ExcErr
Else
    DEVTO = 0
    GoTo Exit_Talk
End If
'************************************************************************
ExcErr:
ExCODE$ = Mid$(RX$, 5, 2)
CODE = Val(ExCODE$)
msg$ = msg & vbCrLf & " Exception Error #" & CODE & " - " & ExErr$(CODE)
msg$ = msg & vbCrLf & "Sent =>" & Send$
msg$ = msg & vbCrLf & "Received =>" & RX$
msg$ = msg & vbCrLf & " Exception Error #" & CODE & " - " & ExErr$(CODE)
MsgBox msg, vbCritical, "Communication Error with PLC"
GoTo Exit_Talk
'************************************************************************
TOCHK:
'CHECKING FOR EXCESSIVE REPLY TIME.
'If so, a retransmission of the message not responded to will occur.
NEWTM! = Timer:
If (NEWTM! - OLDTM!) < 0 Then NEWTM! = NEWTM! + 86400!
If (NEWTM! - OLDTM!) > TMTRY! Then
    Repeat = Repeat + 1
    Fails = Fails + 1
    If Repeat = Maxtry + 1 Then
        GoTo NoReply
    Else
        GoTo SendMess
    End If
End If
GoTo Receeve
'************************************************************************
NoReply:
ComErr = 1
Repeat = 1
msg = "No Reply After" & Maxtry & "Attempts"
MsgBox msg, vbCritical, "Communication Error with PLC"
Repeat = 1
GoTo Exit_Talk
'************************************************************************
CommError:
msg = msg & vbCrLf & " The Message Sent Was = > " & Send$
msg = msg & vbCrLf & "The Reply Received Was => " & RX$
MsgBox msg, vbCritical, "Communication Error with PLC"
EmptyBuffer
GoTo Exit_Talk
'************************************************************************
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Talk "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Talk"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Talk:
'REPLY$ = DLEStr$ + ACKStr$
'frmmain.Comm1.Output = REPLY$
GoTo endT
endT:
CommComp = 1
msg$ = " The Message Sent Was = > " & Send$ 'hexsend$?
msg$ = msg$ & vbCrLf & "The Reply Received Was => " & RX$
'Debug.Print msg$
If Continuous = 0 Then
    If frmMain.Text1.Text <> msg$ & vbCrLf Then
        frmMain.Text1.Text = frmMain.Text1.Text & msg$ & vbCrLf
    End If
Else
    If frmMain.Text1.Text <> msg$ & vbCrLf Then
        frmMain.Text1.Text = msg$ & vbCrLf
    End If
End If
Screen.MousePointer = 0
End Sub

Sub Pick()
'pick off the next character in the reply
If CommInLen = 0 Then RecBuf
CommInLen = Len(CommIn$)
If CommInLen = 0 Then GoTo PICKend
RxChar$ = Left$(CommIn$, 1)
'add the character to the total reply string
RxInTOT$ = RxInTOT$ + RxChar$
CommInLen = CommInLen - 1
If CommInLen = 0 Then
    CommIn$ = ""
    GoTo PICKend
End If
CommIn$ = Right$(CommIn$, CommInLen)
PICKend:
End Sub

Sub RecBuf()
WAITRECS:
'DoEvents
CharInS = frmMain.Comm1.InBufferCount
If CharInS = 0 Then
    'TOCHK ' If INPUT buffer is empty go and check if retry time has passed yet.
    If (NEWTM! - OLDTM!) > TMTRY! Then Exit Sub
    NEWTM! = Timer
    GoTo WAITRECS
End If
frmMain.Comm1.InputLen = CharInS
CommIn$ = frmMain.Comm1.Input
CommIn$ = StrConv(CommIn$, vbUnicode) 'was commin, ... changed 5/98 to CommIn$,
CommInLen = Len(CommIn$)
End Sub

Sub EmptyBuf()
'another way of doing the below routine 1 character at a time
On Error GoTo Oops
Dim RxLeft$
BufLen = frmMain.Comm1.InBufferCount
If BufLen <> 0 Then
    For i = 1 To BufLen
        RxLeft$ = frmMain.Comm1.Input
        RxLeft$ = StrConv(RxLeft$, vbUnicode)
        If RxLeft$ <> "" Then
            frmMain.Caption = frmMain.Caption + Hex$(Asc(RxLeft$))
        End If
    Next i
End If
GoTo Exit_EmptyBuf
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine EmptyBuf "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in EmptyBuf"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_EmptyBuf:
End Sub

Sub EmptyBuffer()
Dim Sinks As Integer
'this just totally empties the input buffer
EmptySink:
'check how many characters are waiting to be read
BufLen = frmMain.Comm1.InBufferCount
'set the input buffer length to the number of characters waiting in the buffer
frmMain.Comm1.InputLen = BufLen
'get all the characters left in the buffer
ERRCHAR$ = frmMain.Comm1.Input
If BufLen > 0 Then
    Debug.Print "Emptying " & BufLen & " Characters..."
End If
'check if there are more characters coming in
If frmMain.Comm1.InBufferCount > 0 Then GoTo EmptySink
For Sinks = 1 To Len(ERRCHAR$)
    frmMain.Caption = frmMain.Caption + Hex$(Asc(Mid$(ERRCHAR$, Sinks, 1)))
Next Sinks
SINKend:
ComErr = 1
If frmMain.Comm1.InBufferCount > 0 Then GoTo EmptySink
'reset our input length to 1 character
frmMain.Comm1.InputLen = 1
End Sub

Sub CRCcheck(MessageIn As String)
'In RTU mode, messages include an error-checking field that is based on a
'Cyclical Redundancy Check (CRC) method. The CRC field checks the contents
'of the entire message. It is applied regardless of any parity check method used
'for the individual characters of the message.
'The CRC field is two bytes, containing a 16-bit binary value. The CRC value is
'calculated by the transmitting device, which appends the CRC to the message.
'The receiving device recalculates a CRC during receipt of the message, and
'compares the calculated value to the actual value it received in the CRC field.
'If the two values are not equal, an error results.
'
'The CRC is started by first preloading a 16-bit register to all 1's.
'Then a process begins of applying successive 8-bit bytes of the message
'to the current contents of the register.
'Only the eight bits of data in each character are used for generating
'the CRC. Start and stop bits, and the parity bit, do not apply to the CRC.
'During generation of the CRC, each 8-bit character is exclusive ORed with the
'register contents.
'Then the result is shifted in the direction of the least significant
'bit (LSB), with a zero filled into the most significant bit (MSB) position. The LSB is
'extracted and examined. If the LSB was a 1, the register is then exclusive ORed
'with a preset, fixed value. If the LSB was a 0, no exclusive OR takes place.
'This process is repeated until eight shifts have been performed. After the last
'(eighth) shift, the next 8-bit byte is exclusive ORed with the register's current
'value, and the process repeats for eight more shifts as described above. The final
'contents of the register, after all the bytes of the message have been applied, is
'the CRC value.
'When the CRC is appended to the message, the low-order byte is appended first,
'followed by the high-order byte.
Dim LoByte As Long
Dim HiByte As Long
Dim CRCconst As Long
Dim CRCreg8 As Long
Dim CRCresult As Long
Dim sl As Integer
Dim DataByte As Integer
Dim OldCRC As Long
Dim DivVal As Integer
Dim CRCHex$
Dim CRCLen As Integer
Dim NLowStr$
Dim NHighStr$
'
On Error GoTo Oops
LoByte = 255
HiByte = Val(Abs(("&h" & "FF00"))) '65280
CRCconst = Val(Abs("&H" & "A001")) '40961
'
CRCReg = Val(Abs(("&h" & "FFFF")))
CRCFrame$ = MessageIn
'rtu sends the 8 bit binary character so we decode each character
beginCRC:
sl = Len(CRCFrame$)
'CRCReg = 0
For i = 1 To sl
    'get the character and its ascii code
    DataByte = Asc(Mid$(CRCFrame$, i, 1))
    'JoyMain.Caption = JoyMain.Caption + Hex$(DataByte) + " "
    CRCRegHi = HiByte And CRCReg
    CRCreg8 = LoByte And CRCReg
    CRCresult = DataByte Xor CRCreg8
    'Xor the data byte and the right 8 bits of the crcreg
    'place the result in the right 8 bits of the crcreg
    CRCRegLo = 255 And CRCresult
    CRCReg = CRCRegHi Or CRCRegLo
    'shift bit right with 0 in on left
    For j = 1 To 8
        OldCRC = CRCReg
        'divide by 2 to do a bit shift right
        'determine value of bit to be shifted out
        DivVal = CRCReg Mod 2
        CRCReg = CRCReg \ 2
        If DivVal = 1 Then
            CRCReg = CRCconst Xor CRCReg
        End If
    Next j
Next i
CRCHex$ = Right$((Hex$(CRCReg)), 4)
CRCHex$ = AddZero(CRCHex$, 4)
CRCLen = Len(CRCHex$)
NLowStr$ = Right$(CRCHex$, 2)
NHighStr$ = Left$(CRCHex$, CRCLen - 2)
CRCRegLo = Val("&H" + NLowStr$)
CRCRegHi = Val("&H" + NHighStr$)
'Debug.Print NLowStr$, NHighStr$
GoTo Exit_CRCcheck
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine CRCcheck "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in CRCcheck"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_CRCcheck:
End Sub

Sub InitModBus()
'SET TO `0' FOR CONTINUOUS TRANSMISSION
TMTRY! = 1
'SET MAXIMUM REPEATS TO FIVE.  THIS CAN BE CHANGED TO ANY INTEGER
Maxtry = 5
HRBase = 0
NodeAddr = 1
End Sub

Function AddZero(StrIn$, intDigits)
'we can actually do this by
'AddZero = Format(StrIn$, String(intDigits, "0"))
'
StrIn$ = Trim(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddZero = StrIn$
    Exit Function
End If
AddZero = String(intDigits - Len(StrIn$), "0") & StrIn$
End Function

Sub Beeep()
'makes a beep on the pc's internal speaker
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
End Sub

Sub Alarm()
'makes an alert beep on the pc's internal speaker
'usage Call NTBeep(CLng(FreqHz), CLng(LengthMs))
NTBeep 100, 50
NTBeep 200, 50
NTBeep 300, 50
NTBeep 400, 50
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
NTBeep 100, 50
NTBeep 200, 50
NTBeep 300, 50
NTBeep 400, 50
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
DoEvents
End Sub

Sub GetIO()
'returns the physical I/O in an array
ReadInputs 1, 100
For i = 1 To 15
    Inputs(i) = ModIn(i)
    Outputs(i) = ModIn(i + 48)
Next i
End Sub
