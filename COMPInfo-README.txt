COMPInfo.js - �������� ��������� ��� ����� ��������� ���������� � ������.

JS-������ �������� �� ������ ���������� ������������� WMI-������� Win32_*. �� �����
������������ ����� �������� ���������, ��������� ������ ��� � �������������� WMI �����
� ������ � ������ ��� �������� ���������, ���������� �� ����, �������� �� ��� ��������
��� ��������� � ����������� ������, �� C#, PShell ��� ���� C++, ������ ���������� ���������� 
���������� �� ����� �������������, ������������� ������ API. ������ ��������� ������ AIDA
�� �������������� ���� ���, ��� �������� ������ WIM-������, �� ���������� ��������������,
�������������� ��������� ������� ��� ���������� ����������.

��������� ��������� ���� COMPInfo, � �������� ������ �������� ��� �������� ���������������� 
���������������, ��� �����, ��� ����� ������� � ���� ���������. ��� ������ � �������� ����.

������ ����� � ������ ���� ����, �������� �������� PShell ��������, ������� ������� ����� �� 
���� ������� ��������������� ����� ���:

1. �������������� ������������ �, ���� ���������, ����� ���� ����������� � ��������� ����.
   � ������ ������ ��� ���������� ������� ��� ��������� ����������� ����� ���� ������������ 
   � ���� Exel - ������ ��� ����� ������ ��������� �������.

2. ���� ��������� ���������� � ��������� ��������� ����������� � ����: ��� �� �������
   ���������� � ������������ �� ���� � ���, ������ � �������� �������� ����� ��� �������
   �������� ���������� ��������������, ������� ���������� ��� ����������.

��� ����� ������ ����� ������� ������ COMPInfo, �������� ���� ���������� � ���, ���
��������� �������������. ������ �� ������� ���������� �� ������������� ������.
��� ������� �����, � ������� ��� �������:

----------------------------------------PhysicalMemory--------------------------
Description:                            Physical Memory
BankLabel:                              Bank 0/1
PositionInRow:                          1
Capacity:                               4.00 GB
DataWidth:                              64 bits
FormFactor:                             12 (SODIMM)
InterleaveDataDepth:                    0
InterleavePosition:                     0
Manufacturer:                           80AD (Hynix)
PartNumber:                             HMT351S6BFR8C-PB
SerialNumber:                           111489FC
MemoryType:                             24 (DDR3)
Speed:                                  1066
TypeDetail:                             128 (Synchronous)

�������� �������������, �������� ���� 80AD, ����� ��������� � ����� � ������� � �������
� ���� ��� ����������� "80AD" --> Hynix. ( ������ ���� � ������ ������ �������������� ����
������� ���� 80AD ). �� ��� ������ ������:

----------------------------------------PhysicalMemory--------------------------
Description:                            Physical Memory
BankLabel:                              BANK 0
PositionInRow:                          1
Capacity:                               4.00 GB
ConfiguredClockSpeed:                   1600
DataWidth:                              64 bits
FormFactor:                             12 (SODIMM)
InterleaveDataDepth:                    0
InterleavePosition:                     0
Manufacturer:                           0000            <========
PartNumber:                             SHARETRONIC
SerialNumber:                           00000000        <========
MemoryType:                             24 (DDR3)
Speed:                                  1600
TypeDetail:                             128 (Synchronous)

����� �����-�� �� �������������� ������ ����� ������ ������ ���������� � ������
������ ������������� ����� ����� ��� ������ ��������������� ���� � ���������, ���
�� �������� �� SODIMM. ���� ���� ������� ������� ����� �������������� ���������
� ���, ��� ����� ���������� � ����� ������ � ����� ������� ������� ������� ��� �����
�������� ������.

��� ������ ��� ������ �������� �������������, ������� ������ �������� ��� ������������
����, �� ��� ��������� � ����� ������������� � ����� "00000000000000CE" ��� ���� �� �������.

Description:                            Physical Memory
BankLabel:                              Bank 00
PositionInRow:                          1
Capacity:                               1.00 GB
DataWidth:                              64 bits
FormFactor:                             12 (SODIMM)
InterleaveDataDepth:                    0
InterleavePosition:                     0
Manufacturer:                           00000000000000CE
PartNumber:                             M4 70T2864QZ3-CF7
SerialNumber:                           CE00000000000000010833483A5667
MemoryType:                             21 (DDR2)
Speed:                                  667
TypeDetail:                             128 (Synchronous)

������� � ���� � ��� ������� �������: �� ���� ��������� ������, ��� �� ������, �� ������
������������� ����� � ����� ������. � ���������� ������ ����-���� RAMInfo.cmd, ������� 
��������� ��� �� ����� COMPInfo.js, �� ����������� ���������� ���� � ���������� ������
� ������� ������ ����� ��������� �������� - ���� �������, ����������� ���������������
��� ������� �� DIMM, SODIMM, ������������� � ����� �����. �������� �� ���������� �������, 
���������� � ��������, �� �� ����� ��� ������, �������� � ���� �� ���� ��������� �������������
� ��� ����� ������ ��� ������� � ��������� ������� ������������ ����� � �������� ����.

COMPInfo.cmd ������ ������ ���������� � ���� ���������� �����, ������������� COMPInfo.js.
������ ����� ������������ ������� ���� ����������, ������� � ��������� ������ ��������
����������, ������������ ����� ������� ��������� ������� � ���. ��� ������ ���������� ������
� ���������� ������, � ���� ������ ����������� ����-���� RAMInfo.cmd, �� ����� ���� �� 
�������� � ��� ����, �������� "COMPInfo /select:ram". ������ SELECT ������������ ��� ������������
����� ������� ��������� ���������: PLATFORM, CPU, RAM, DISKS, VIDEO, SOUND, BATTERY. 
���� ������� ���, �� �������� ��� ������� + ��� ���-���. ����� ����, ��������� ���������
����� ��������� ���������� ��������.
 




 



