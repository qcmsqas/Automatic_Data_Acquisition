Attribute VB_Name = "FunctionDeclare"
Option Explicit

Declare Function auto_set Lib "MPC08" () As Long
Declare Function init_board Lib "MPC08" () As Long
Declare Function get_max_axe Lib "MPC08" () As Long
Declare Function get_board_num Lib "MPC08" () As Long
Declare Function get_axe Lib "MPC08" (ByVal board_no As Long) As Long
Declare Function set_outmode Lib "MPC08" (ByVal ch As Long, ByVal mode As Long, ByVal logic As Long) As Long
Declare Function set_home_mode Lib "MPC08" (ByVal ch As Long, ByVal origin_mode As Long) As Long
Declare Function set_maxspeed Lib "MPC08" (ByVal ch As Long, ByVal maxspeed As Double) As Long
Declare Function set_conspeed Lib "MPC08" (ByVal ch As Long, ByVal conspeed As Double) As Long
Declare Function get_conspeed Lib "MPC08" (ByVal ch As Long) As Double
Declare Function set_profile Lib "MPC08" (ByVal ch As Long, ByVal vl As Double, ByVal vh As Double, ByVal ad As Double) As Long
Declare Function get_profile Lib "MPC08" (ByVal ch As Long, ByRef vl As Double, ByRef vh As Double, ByRef ad As Double) As Long
Declare Function set_vector_conspeed Lib "MPC08" (ByVal conspeed As Double) As Long
Declare Function set_vector_profile Lib "MPC08" (ByVal vec_vl As Double, ByVal vec_vh As Double, ByVal vec_ad As Double) As Long
Declare Function get_vector_conspeed Lib "MPC08" () As Double
Declare Function get_vector_profile Lib "MPC08" (ByRef vec_vl As Double, ByRef vec_vh As Double, ByRef vec_ad As Double) As Long
Declare Function get_rate Lib "MPC08" (ByVal ch As Long) As Double

'运动指令函数
Declare Function con_pmove Lib "MPC08" (ByVal ch As Long, ByVal step As Long) As Long
Declare Function fast_pmove Lib "MPC08" (ByVal ch As Long, ByVal step As Long) As Long
Declare Function con_pmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long) As Long
Declare Function fast_pmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long) As Long
Declare Function con_pmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long, ByVal ch3 As Long, ByVal step3 As Long) As Long
Declare Function fast_pmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long, ByVal ch3 As Long, ByVal step3 As Long) As Long
Declare Function con_pmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long, ByVal ch3 As Long, ByVal step3 As Long, ByVal ch4 As Long, ByVal step4 As Long) As Long
Declare Function fast_pmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal step1 As Long, ByVal ch2 As Long, ByVal step2 As Long, ByVal ch3 As Long, ByVal step3 As Long, ByVal ch4 As Long, ByVal step4 As Long) As Long
Declare Function con_vmove Lib "MPC08" (ByVal ch As Long, ByVal dir As Long) As Long
Declare Function fast_vmove Lib "MPC08" (ByVal ch As Long, ByVal dir As Long) As Long
Declare Function con_vmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long) As Long
Declare Function fast_vmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long) As Long
Declare Function con_vmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long) As Long
Declare Function fast_vmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long) As Long
Declare Function con_vmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long, ByVal ch4 As Long, ByVal dir4 As Long) As Long
Declare Function fast_vmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long, ByVal ch4 As Long, ByVal dir4 As Long) As Long
Declare Function con_hmove Lib "MPC08" (ByVal ch As Long, ByVal dir As Long) As Long
Declare Function fast_hmove Lib "MPC08" (ByVal ch As Long, ByVal dir As Long) As Long
Declare Function con_hmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long) As Long
Declare Function fast_hmove2 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long) As Long
Declare Function con_hmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long) As Long
Declare Function fast_hmove3 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long) As Long
Declare Function con_hmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long, ByVal ch4 As Long, ByVal dir4 As Long) As Long
Declare Function fast_hmove4 Lib "MPC08" (ByVal ch1 As Long, ByVal dir1 As Long, ByVal ch2 As Long, ByVal dir2 As Long, ByVal ch3 As Long, ByVal dir3 As Long, ByVal ch4 As Long, ByVal dir4 As Long) As Long
Declare Function con_line2 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long) As Long
Declare Function con_line3 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long, ByVal ch3 As Long, ByVal pos3 As Long) As Long
Declare Function con_line4 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long, ByVal ch3 As Long, ByVal pos3 As Long, ByVal ch4 As Long, ByVal pos4 As Long) As Long
Declare Function fast_line2 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long) As Long
Declare Function fast_line3 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long, ByVal ch3 As Long, ByVal pos3 As Long) As Long
Declare Function fast_line4 Lib "MPC08" (ByVal ch1 As Long, ByVal pos1 As Long, ByVal ch2 As Long, ByVal pos2 As Long, ByVal ch3 As Long, ByVal pos3 As Long, ByVal ch4 As Long, ByVal pos4 As Long) As Long
Declare Function change_pos Lib "MPC08" (ByVal ch As Long, ByVal pos As Long) As Long

'制动函数
Declare Function sudden_stop Lib "MPC08" (ByVal ch As Long) As Long
Declare Function sudden_stop2 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long) As Long
Declare Function sudden_stop3 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long, ByVal ch3 As Long) As Long
Declare Function sudden_stop4 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long, ByVal ch3 As Long, ByVal ch4 As Long) As Long
Declare Function decel_stop Lib "MPC08" (ByVal ch As Long) As Long
Declare Function decel_stop2 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long) As Long
Declare Function decel_stop3 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long, ByVal ch3 As Long) As Long
Declare Function decel_stop4 Lib "MPC08" (ByVal ch1 As Long, ByVal ch2 As Long, ByVal ch3 As Long, ByVal ch4 As Long) As Long

'位置和状态设置函数
Declare Function set_abs_pos Lib "MPC08" (ByVal ch As Long, ByVal pos As Long) As Long
Declare Function reset_pos Lib "MPC08" (ByVal ch As Long) As Long
Declare Function reset_enc_pos Lib "MPC08" (ByVal ch As Long) As Long
Declare Function reset_cmd_counter Lib "MPC08" () As Long
Declare Function set_dir Lib "MPC08" (ByVal ch As Long, ByVal dir As Long) As Long
Declare Function enable_sd Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function enable_el Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function enable_org Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function set_sd_logic Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function set_el_logic Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function set_org_logic Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long
Declare Function set_alm_logic Lib "MPC08" (ByVal ch As Long, ByVal flag As Long) As Long

'状态查询函数
Declare Function get_abs_pos Lib "MPC08" (ByVal ch As Long, ByRef pos As Long) As Long
Declare Function get_rel_pos Lib "MPC08" (ByVal ch As Long, ByRef pos As Long) As Long
Declare Function get_cur_dir Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_status Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_done Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_limit Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_home Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_SD Lib "MPC08" (ByVal ch As Long) As Long
Declare Function check_alarm Lib "MPC08" (ByVal ch As Long) As Long
Declare Function get_cmd_counter Lib "MPC08" () As Long

'I/O口操作函数
Declare Function checkin_byte Lib "MPC08" (ByVal cardno As Long) As Long
Declare Function checkin_bit Lib "MPC08" (ByVal cardno As Long, ByVal bitno As Long) As Long
Declare Function outport_bit Lib "MPC08" (ByVal cardno As Long, ByVal bitno As Long, ByVal status As Long) As Long
Declare Function outport_byte Lib "MPC08" (ByVal cardno As Long, ByVal bytedata As Long) As Long
Declare Function check_SFR Lib "MPC08" (ByVal cardno As Long) As Long

'其它函数
Declare Function set_backlash Lib "MPC08" (ByVal axis As Long, ByVal blash As Long) As Long
Declare Function start_backlash Lib "MPC08" (ByVal axis As Long) As Long
Declare Function end_backlash Lib "MPC08" (ByVal axis As Long) As Long
Declare Function change_speed Lib "MPC08" (ByVal ch As Long, ByVal speed As Double) As Long
Declare Function change_accel Lib "MPC08" (ByVal ch As Long, ByVal accel As Double) As Long
Declare Function Outport Lib "MPC08" (ByVal PortID As Long, ByVal OutPortData As Byte) As Long
Declare Function Inport Lib "MPC08" (ByVal PortID As Long) As Long
Declare Function set_ramp_flag Lib "MPC08" (ByVal flag As Long) As Long
Declare Function get_lib_ver Lib "MPC08" (ByRef major As Long, ByRef minor1 As Long, ByRef minor2 As Long) As Long
Declare Function get_sys_ver Lib "MPC08" (ByRef major As Long, ByRef minor1 As Long, ByRef minor2 As Long) As Long
Declare Function get_card_ver Lib "MPC08" (ByVal cardno As Long, ByRef cardtype As Long, ByRef major As Long, ByRef minor1 As Long, ByRef minor2 As Long) As Long
Declare Function get_cardno Lib "MPC08" (ByRef cardno1 As Long, ByRef cardno2 As Long, ByRef cardno3 As Long, ByRef cardno4 As Long) As Long

'编码器操作函数
Declare Function set_getpos_mode Lib "MPC08" (ByVal ch As Long, ByVal mode As Long) As Long
Declare Function set_encoder_mode Lib "MPC08" (ByVal ch As Long, ByVal mode As Long, ByVal multip As Long, ByVal count_unit As Long) As Long
Declare Function get_encoder Lib "MPC08" (ByVal ch As Long, ByRef pos As Long) As Long

'同步位置输出控制
Declare Function enable_io_pos Lib "MPC08" (ByVal cardno As Long, ByVal flag As Long) As Long
Declare Function set_io_pos Lib "MPC08" (ByVal ch As Long, ByVal open_pos As Long, ByVal close_pos As Long) As Long

'错误代码获取
Declare Function get_last_err Lib "MPC08" () As Long
Declare Function get_err Lib "MPC08" (ByVal index As Long, ByRef data As Long) As Long
Declare Function reset_err Lib "MPC08" () As Long

