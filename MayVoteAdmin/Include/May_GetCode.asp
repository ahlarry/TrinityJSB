<%
Option Explicit
Response.buffer=true
Call Com_CreatValidCode("GetCode")

Sub Com_CreatValidCode(pSN)
	'Author: Layen
	'QQ: 84815733
	'E-mail: support@ssaw.net
	' ��ֹ����
	Response.Expires = -9999 
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	Response.ContentType = "Image/BMP"
	
	Randomize
	
	Dim i, ii, iii
	
	Const cOdds = 0 ' �ӵ���ֵĻ���
	Const cAmount = 8 ' ��������
	Const cCode = "ABCDEFGH"
	
	' ��ɫ������(�ַ�������)
	Dim vColorData(1)
	vColorData(0) = ChrB(0) & ChrB(0) & ChrB(0)  ' ��0����0����0����ɫ��
	vColorData(1) = ChrB(255) & ChrB(255) & ChrB(255) ' ��250����236����211��ǳ��ɫ��
	
	' ��������ַ�
	Dim vCode(4), vCodes
	For i = 0 To 3
	  vCode(i) = Int(Rnd * cAmount)
	  vCodes = vCodes & Mid(cCode, vCode(i) + 1, 1)
	Next
	Session(pSN) = vCodes  '��¼��Session
	' �ַ�������
	Dim vNumberData(9)
	vNumberData(0) = "1111100111111110101111111010011111011101111100000111110111011110011101111111110111111111011111111111"
	vNumberData(1) = "1111111111111000001111110110111111011011111100001111110110111111011011111101101111110110111110000011"
	vNumberData(2) = "1111111111111000000111011111111001111111101111111110111111111011111111110011111111100000011111111111"
	vNumberData(3) = "0000000000001111110000100001100010000010001000001000100000100010000010001000011000111111000000000000"
	vNumberData(4) = "0000000000011111111001000000000100000000011111000001000000000100000000011111111000000000000000000000"
	vNumberData(5) = "0000000000001111111000100000000010000000001111111000100000000010000000001000000000100000000000000000"
	vNumberData(6) = "1111111111111000001111011111111001111111101111111110100000011011111011100111101111000110111111000011"
	vNumberData(7) = "0000000000010000001001000000100100000010011111111001000000100100000010010000001000000000000000000000"
	' ���ͼ���ļ�ͷ
	Response.BinaryWrite ChrB(66) & ChrB(77) & ChrB(230) & ChrB(4) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(10) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0)
	
	' ���ͼ����Ϣͷ
	Response.BinaryWrite ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(176) & ChrB(4) &_
	  ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) &_
	  ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0)
	
	For i = 9 To 0 Step -1  ' ����������
	  For ii = 0 To 3  ' ����������
	   For iii = 1 To 10 ' ������������
	    ' ���С����֡������ص����ͼ������
	    If Rnd * 99 + 1 < cOdds Then ' ��������ӵ�
	     Response.BinaryWrite vColorData(0)
	    Else
	     Response.BinaryWrite vColorData(Mid(vNumberData(vCode(ii)), i * 10 + iii, 1))
	    End If
	   Next
	  Next
	Next
End Sub
%>
