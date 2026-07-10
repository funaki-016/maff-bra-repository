''' <summary>
''' メッセージID
''' </summary>
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前            |                修  正  内  容
'// -----------+------------+-------------------------+---------------------------------------------
'//  REV_001   | 2020.10.26 |TSP)                     | フェーズ3 要件No.2修正
'//  REV_002   | 2021.12.16 |日本コンピュータシステム | 要件10 追加
'//  REV_003   | 2021.12.20 |日本コンピュータシステム | 要件19 追加
'//  REV_004   | 2021.12.23 |日本コンピュータシステム | 要件11 追加
'//  REV_005   | 2021.12.17 |日本コンピュータシステム | 要件1-⑤ 追加
'//  REV_006   | 2021.12.27 |日本コンピュータシステム | 要件2 追加
'//  REV_007   | 2022.10.06 |大興電子通信             | 要件No.20 追加
'//  REV_008   | 2022.12.23 |大興電子通信             | 要件No.15 追加
'//  REV_009   | 2023.01.13 |大興電子通信             | 要件No.5 追加
'//  REV_010   | 2023.10.30 |大興電子通信             | 要件No.1 追加
'//  REV_011   | 2023.11.27 |大興電子通信             | 要件No.20 追加
'//  REV_012   | 2024.05.31 |大興電子通信             | 要件No.1 追加
'//*************************************************************************************************
''' <remarks></remarks>
Public Class MessageID

    ''' <summary>メッセージID：確認(Q)</summary>
    Public Const MSG_Q_001 As String = "BRAQ-001"
    Public Const MSG_Q_002 As String = "BRAQ-002"
    Public Const MSG_Q_003 As String = "BRAQ-003"
    Public Const MSG_Q_004 As String = "BRAQ-004"
    Public Const MSG_Q_005 As String = "BRAQ-005"
    Public Const MSG_Q_006 As String = "BRAQ-006"
    Public Const MSG_Q_007 As String = "BRAQ-007"
    Public Const MSG_Q_008 As String = "BRAQ-008"
    Public Const MSG_Q_009 As String = "BRAQ-009"
    Public Const MSG_Q_010 As String = "BRAQ-010"
    Public Const MSG_Q_011 As String = "BRAQ-011"
    Public Const MSG_Q_012 As String = "BRAQ-012"
    Public Const MSG_Q_013 As String = "BRAQ-013"
    Public Const MSG_Q_014 As String = "BRAQ-014"
    Public Const MSG_Q_015 As String = "BRAQ-015"
    Public Const MSG_Q_016 As String = "BRAQ-016"
    Public Const MSG_Q_017 As String = "BRAQ-017"
    Public Const MSG_Q_018 As String = "BRAQ-018"
    Public Const MSG_Q_019 As String = "BRAQ-019"
    Public Const MSG_Q_020 As String = "BRAQ-020"
    Public Const MSG_Q_021 As String = "BRAQ-021"
    Public Const MSG_Q_022 As String = "BRAQ-022"
    Public Const MSG_Q_023 As String = "BRAQ-023"
    Public Const MSG_Q_024 As String = "BRAQ-024"
    Public Const MSG_Q_025 As String = "BRAQ-025"
    Public Const MSG_Q_026 As String = "BRAQ-026"
    Public Const MSG_Q_027 As String = "BRAQ-027"
    Public Const MSG_Q_028 As String = "BRAQ-028"
    Public Const MSG_Q_029 As String = "BRAQ-029"
    Public Const MSG_Q_030 As String = "BRAQ-030"
    Public Const MSG_Q_031 As String = "BRAQ-031"
    Public Const MSG_Q_032 As String = "BRAQ-032"
    Public Const MSG_Q_033 As String = "BRAQ-033"
    Public Const MSG_Q_034 As String = "BRAQ-034"
    '--- REV_001 ADD START
    Public Const MSG_Q_035 As String = "BRAQ-035"
    Public Const MSG_Q_036 As String = "BRAQ-036"
    Public Const MSG_Q_037 As String = "BRAQ-037"
    Public Const MSG_Q_038 As String = "BRAQ-038"
    Public Const MSG_Q_039 As String = "BRAQ-039"
    Public Const MSG_Q_040 As String = "BRAQ-040"
    Public Const MSG_Q_041 As String = "BRAQ-041"
    Public Const MSG_Q_042 As String = "BRAQ-042"
    Public Const MSG_Q_043 As String = "BRAQ-043"
    Public Const MSG_Q_044 As String = "BRAQ-044"
    Public Const MSG_Q_045 As String = "BRAQ-045"
    Public Const MSG_Q_046 As String = "BRAQ-046"
    Public Const MSG_Q_047 As String = "BRAQ-047"
    Public Const MSG_Q_048 As String = "BRAQ-048"
    '--- REV_001 ADD END
    '--- REV 005 START
    Public Const MSG_Q_050 As String = "BRAQ-050"
    '--- REV 005 END
    '--- REV_003 ADD START
    Public Const MSG_Q_051 As String = "BRAQ-051"
    '--- REV_003 ADD END
    '--- REV_002 ADD START
    Public Const MSG_Q_052 As String = "BRAQ-052"
    '--- REV_002 ADD END
    Public Const MSG_Q_053 As String = "BRAQ-053" 'REV_006 ADD
    Public Const MSG_Q_054 As String = "BRAQ-054" 'REV_006 ADD
    Public Const MSG_Q_055 As String = "BRAQ-055" 'REV_006 ADD
    Public Const MSG_Q_056 As String = "BRAQ-056" '2022/03/10 ADD 連絡票No303対応
    ' REV_009↓
    Public Const MSG_Q_057 As String = "BRAQ-057"
    Public Const MSG_Q_058 As String = "BRAQ-058"
    Public Const MSG_Q_059 As String = "BRAQ-059"
    Public Const MSG_Q_060 As String = "BRAQ-060"
    Public Const MSG_Q_061 As String = "BRAQ-061"
    Public Const MSG_Q_062 As String = "BRAQ-062"
    Public Const MSG_Q_063 As String = "BRAQ-063"
    Public Const MSG_Q_064 As String = "BRAQ-064"
    ' REV_009↑
    ' REV_010↓
    Public Const MSG_Q_065 As String = "BRAQ-065"
    Public Const MSG_Q_066 As String = "BRAQ-066"
    Public Const MSG_Q_067 As String = "BRAQ-067"
    Public Const MSG_Q_068 As String = "BRAQ-068"
    ' REV_010↑
    'REV_012 ADD START
    Public Const MSG_Q_069 As String = "BRAQ-069"
    'REV_012 ADD END

    ''' <summary>メッセージID：通知(I)</summary>
    Public Const MSG_I_001 As String = "BRAI-001"
    Public Const MSG_I_002 As String = "BRAI-002"
    Public Const MSG_I_003 As String = "BRAI-003"
    Public Const MSG_I_004 As String = "BRAI-004"
    Public Const MSG_I_005 As String = "BRAI-005"
    Public Const MSG_I_006 As String = "BRAI-006"
    Public Const MSG_I_007 As String = "BRAI-007"
    Public Const MSG_I_008 As String = "BRAI-008"
    Public Const MSG_I_009 As String = "BRAI-009"
    Public Const MSG_I_010 As String = "BRAI-010"
    Public Const MSG_I_011 As String = "BRAI-011"
    Public Const MSG_I_012 As String = "BRAI-012"
    Public Const MSG_I_013 As String = "BRAI-013"
    Public Const MSG_I_014 As String = "BRAI-014"
    Public Const MSG_I_015 As String = "BRAI-015"
    Public Const MSG_I_016 As String = "BRAI-016"
    Public Const MSG_I_017 As String = "BRAI-017"
    Public Const MSG_I_018 As String = "BRAI-018"
    Public Const MSG_I_019 As String = "BRAI-019"
    Public Const MSG_I_020 As String = "BRAI-020"
    Public Const MSG_I_021 As String = "BRAI-021"
    Public Const MSG_I_022 As String = "BRAI-022"
    Public Const MSG_I_023 As String = "BRAI-023"
    Public Const MSG_I_024 As String = "BRAI-024"
    Public Const MSG_I_025 As String = "BRAI-025"
    Public Const MSG_I_026 As String = "BRAI-026"
    Public Const MSG_I_027 As String = "BRAI-027"
    Public Const MSG_I_028 As String = "BRAI-028"
    Public Const MSG_I_029 As String = "BRAI-029"
    Public Const MSG_I_030 As String = "BRAI-030"
    Public Const MSG_I_031 As String = "BRAI-031"
    Public Const MSG_I_032 As String = "BRAI-032"
    Public Const MSG_I_033 As String = "BRAI-033"
    Public Const MSG_I_034 As String = "BRAI-034"
    Public Const MSG_I_035 As String = "BRAI-035"
    '--- REV_001 ADD START
    Public Const MSG_I_036 As String = "BRAI-036"
    Public Const MSG_I_037 As String = "BRAI-037"
    Public Const MSG_I_038 As String = "BRAI-038"
    Public Const MSG_I_039 As String = "BRAI-039"
    Public Const MSG_I_040 As String = "BRAI-040"
    Public Const MSG_I_041 As String = "BRAI-041"
    Public Const MSG_I_042 As String = "BRAI-042"
    Public Const MSG_I_043 As String = "BRAI-043"
    Public Const MSG_I_044 As String = "BRAI-044"
    Public Const MSG_I_045 As String = "BRAI-045"
    '--- REV_001 ADD END
    '--- REV 005 START
    Public Const MSG_I_047 As String = "BRAI-047"
    '--- REV 005 END
    Public Const MSG_I_049 As String = "BRAI-049" 'REV_006 ADD
    ' REV_009↓
    Public Const MSG_I_050 As String = "BRAI-050"
    Public Const MSG_I_051 As String = "BRAI-051"
    Public Const MSG_I_052 As String = "BRAI-052"
    Public Const MSG_I_053 As String = "BRAI-053"
    ' REV_009↑
    ' REV_010↓
    Public Const MSG_I_054 As String = "BRAI-054"
    Public Const MSG_I_055 As String = "BRAI-055"
    ' REV_010↑
    ' REV_011↓
    Public Const MSG_I_056 As String = "BRAI-056"
    ' REV_011↑

    ''' <summary>メッセージID：警告(W)</summary>
    ' REV_010↓
    Public Const MSG_W_001 As String = "BRAW-001"
    ' REV_010↑
    ' REV_011↓
    Public Const MSG_W_002 As String = "BRAW-002"
    ' REV_011↑

    ''' <summary>メッセージID：エラー(E)</summary>
    Public Const MSG_E_001 As String = "BRAE-001"
    Public Const MSG_E_002 As String = "BRAE-002"
    Public Const MSG_E_003 As String = "BRAE-003"
    Public Const MSG_E_004 As String = "BRAE-004"
    Public Const MSG_E_005 As String = "BRAE-005"
    Public Const MSG_E_006 As String = "BRAE-006"
    Public Const MSG_E_007 As String = "BRAE-007"
    Public Const MSG_E_008 As String = "BRAE-008"
    Public Const MSG_E_009 As String = "BRAE-009"
    Public Const MSG_E_010 As String = "BRAE-010"
    Public Const MSG_E_011 As String = "BRAE-011"
    Public Const MSG_E_012 As String = "BRAE-012"
    Public Const MSG_E_013 As String = "BRAE-013"
    Public Const MSG_E_014 As String = "BRAE-014"
    Public Const MSG_E_015 As String = "BRAE-015"
    Public Const MSG_E_016 As String = "BRAE-016"
    Public Const MSG_E_017 As String = "BRAE-017"
    Public Const MSG_E_018 As String = "BRAE-018"
    Public Const MSG_E_019 As String = "BRAE-019"
    Public Const MSG_E_020 As String = "BRAE-020"
    Public Const MSG_E_021 As String = "BRAE-021"
    Public Const MSG_E_022 As String = "BRAE-022"
    Public Const MSG_E_023 As String = "BRAE-023"
    Public Const MSG_E_024 As String = "BRAE-024"
    Public Const MSG_E_025 As String = "BRAE-025"
    Public Const MSG_E_026 As String = "BRAE-026"
    Public Const MSG_E_029 As String = "BRAE-029"
    Public Const MSG_E_030 As String = "BRAE-030"
    Public Const MSG_E_031 As String = "BRAE-031"
    Public Const MSG_E_032 As String = "BRAE-032"
    Public Const MSG_E_033 As String = "BRAE-033"
    Public Const MSG_E_034 As String = "BRAE-034"
    Public Const MSG_E_035 As String = "BRAE-035"
    Public Const MSG_E_036 As String = "BRAE-036"
    Public Const MSG_E_037 As String = "BRAE-037"
    Public Const MSG_E_038 As String = "BRAE-038"
    Public Const MSG_E_039 As String = "BRAE-039"
    Public Const MSG_E_040 As String = "BRAE-040"
    Public Const MSG_E_041 As String = "BRAE-041"
    Public Const MSG_E_042 As String = "BRAE-042"
    Public Const MSG_E_043 As String = "BRAE-043"
    Public Const MSG_E_044 As String = "BRAE-044"
    Public Const MSG_E_045 As String = "BRAE-045"
    Public Const MSG_E_046 As String = "BRAE-046"
    Public Const MSG_E_047 As String = "BRAE-047"
    Public Const MSG_E_048 As String = "BRAE-048"
    Public Const MSG_E_049 As String = "BRAE-049"
    Public Const MSG_E_050 As String = "BRAE-050"
    Public Const MSG_E_051 As String = "BRAE-051"
    Public Const MSG_E_052 As String = "BRAE-052"
    Public Const MSG_E_053 As String = "BRAE-053"
    Public Const MSG_E_054 As String = "BRAE-054"
    Public Const MSG_E_055 As String = "BRAE-055"
    Public Const MSG_E_056 As String = "BRAE-056"
    Public Const MSG_E_057 As String = "BRAE-057"
    Public Const MSG_E_058 As String = "BRAE-058"
    Public Const MSG_E_059 As String = "BRAE-059"
    Public Const MSG_E_060 As String = "BRAE-060"
    Public Const MSG_E_061 As String = "BRAE-061"
    Public Const MSG_E_062 As String = "BRAE-062"
    Public Const MSG_E_063 As String = "BRAE-063"
    Public Const MSG_E_064 As String = "BRAE-064"
    Public Const MSG_E_065 As String = "BRAE-065"
    Public Const MSG_E_066 As String = "BRAE-066"
    Public Const MSG_E_067 As String = "BRAE-067"
    '--- REV_001 ADD START
    Public Const MSG_E_068 As String = "BRAE-068"
    Public Const MSG_E_069 As String = "BRAE-069"
    Public Const MSG_E_070 As String = "BRAE-070"
    Public Const MSG_E_071 As String = "BRAE-071"
    Public Const MSG_E_072 As String = "BRAE-072"
    Public Const MSG_E_073 As String = "BRAE-073"
    Public Const MSG_E_074 As String = "BRAE-074"
    Public Const MSG_E_075 As String = "BRAE-075"
    Public Const MSG_E_076 As String = "BRAE-076"
    Public Const MSG_E_077 As String = "BRAE-077"
    Public Const MSG_E_078 As String = "BRAE-078"
    Public Const MSG_E_079 As String = "BRAE-079"
    Public Const MSG_E_080 As String = "BRAE-080"
    Public Const MSG_E_081 As String = "BRAE-081"
    Public Const MSG_E_082 As String = "BRAE-082"
    Public Const MSG_E_083 As String = "BRAE-083"
    Public Const MSG_E_084 As String = "BRAE-084"
    Public Const MSG_E_085 As String = "BRAE-085"
    Public Const MSG_E_086 As String = "BRAE-086"
    Public Const MSG_E_087 As String = "BRAE-087"
    Public Const MSG_E_088 As String = "BRAE-088"
    Public Const MSG_E_089 As String = "BRAE-089"
    Public Const MSG_E_090 As String = "BRAE-090"
    Public Const MSG_E_091 As String = "BRAE-091"
    Public Const MSG_E_092 As String = "BRAE-092"
    Public Const MSG_E_093 As String = "BRAE-093"
    Public Const MSG_E_094 As String = "BRAE-094"
    Public Const MSG_E_095 As String = "BRAE-095"
    Public Const MSG_E_096 As String = "BRAE-096"
    Public Const MSG_E_097 As String = "BRAE-097"
    Public Const MSG_E_098 As String = "BRAE-098"
    '--- REV_001 ADD END
    Public Const MSG_E_102 As String = "BRAE-102" 'REV_005 ADD
    Public Const MSG_E_103 As String = "BRAE-103" 'REV_005 ADD
    Public Const MSG_E_104 As String = "BRAE-104" 'REV_002 ADD
    Public Const MSG_E_105 As String = "BRAE-105" 'REV_004 ADD
    Public Const MSG_E_106 As String = "BRAE-106" 'REV_006 ADD
    Public Const MSG_E_107 As String = "BRAE-107" 'REV_006 ADD
    Public Const MSG_E_108 As String = "BRAE-108" 'REV_006 ADD
    Public Const MSG_E_109 As String = "BRAE-109" 'REV_006 ADD
    Public Const MSG_E_110 As String = "BRAE-110" 'REV_006 ADD
    Public Const MSG_E_112 As String = "BRAE-112" 'REV_004 ADD
    Public Const MSG_E_113 As String = "BRAE-113" 'REV_002 ADD
    Public Const MSG_E_114 As String = "BRAE-114" 'REV_002 ADD
    Public Const MSG_E_115 As String = "BRAE-115" 'REV_004 ADD
    Public Const MSG_E_116 As String = "BRAE-116" 'REV_002 ADD
    ' REV_007↓
    Public Const MSG_E_117 = "BRAE-117"
    Public Const MSG_E_118 = "BRAE-118"
    ' REV_007↑
    ' REV_009↓
    Public Const MSG_E_119 = "BRAE-119"
    Public Const MSG_E_120 = "BRAE-120"
    Public Const MSG_E_121 = "BRAE-121"
    Public Const MSG_E_122 = "BRAE-122"
    Public Const MSG_E_123 = "BRAE-123"
    Public Const MSG_E_124 = "BRAE-124"
    Public Const MSG_E_125 = "BRAE-125"
    Public Const MSG_E_126 = "BRAE-126"
    Public Const MSG_E_127 = "BRAE-127"
    Public Const MSG_E_128 = "BRAE-128"
    Public Const MSG_E_129 = "BRAE-129"
    Public Const MSG_E_130 = "BRAE-130"
    Public Const MSG_E_131 = "BRAE-131"
    Public Const MSG_E_132 = "BRAE-132"
    Public Const MSG_E_133 = "BRAE-133"
    Public Const MSG_E_134 = "BRAE-134"
    Public Const MSG_E_135 = "BRAE-135"
    Public Const MSG_E_136 = "BRAE-136"
    Public Const MSG_E_137 = "BRAE-137"
    Public Const MSG_E_138 = "BRAE-138"
    ' REV_008↓
    Public Const MSG_E_139 = "BRAE-139"
    ' REV_008↑
    Public Const MSG_E_140 = "BRAE-140"
    ' REV_009↑
    ' REV_010↓
    Public Const MSG_E_141 = "BRAE-141"
    Public Const MSG_E_142 = "BRAE-142"
    Public Const MSG_E_143 = "BRAE-143"
    Public Const MSG_E_144 = "BRAE-144"
    Public Const MSG_E_145 = "BRAE-145"
    Public Const MSG_E_146 = "BRAE-146"
    Public Const MSG_E_147 = "BRAE-147"
    Public Const MSG_E_148 = "BRAE-148"
    Public Const MSG_E_149 = "BRAE-149"
    Public Const MSG_E_150 = "BRAE-150"
    Public Const MSG_E_151 = "BRAE-151"
    Public Const MSG_E_152 = "BRAE-152"
    Public Const MSG_E_153 = "BRAE-153"
    Public Const MSG_E_154 = "BRAE-154"
    Public Const MSG_E_155 = "BRAE-155"
    Public Const MSG_E_156 = "BRAE-156"
    Public Const MSG_E_157 = "BRAE-157"
    Public Const MSG_E_158 = "BRAE-158"
    Public Const MSG_E_159 = "BRAE-159"
    ' REV_010↑
    ' REV_011↓
    Public Const MSG_E_160 As String = "BRAE-160"
    ' REV_011↑
    'REV_012 ADD START
    Public Const MSG_E_161 As String = "BRAE-161"
    'REV_012 ADD END

    Public Const MSG_E_999 As String = "BRAE-999"

End Class
