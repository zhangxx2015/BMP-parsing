VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Any Bitmap Reader"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox CheckDBB 
      Caption         =   "use double buffer"
      Height          =   405
      Left            =   1650
      TabIndex        =   2
      Top             =   4920
      Width           =   2265
   End
   Begin VB.ComboBox ComboBMPS 
      Height          =   300
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4950
      Width           =   1275
   End
   Begin VB.PictureBox PCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   150
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   0
      Top             =   120
      Width           =   6240
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)


Private Declare Sub CopyMemoryLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal DestinationPtr As Long, ByVal SourcePtr As Long, ByVal Length As Long)

'BMP文件头
Private Type BITMAPFILEHEADER
        bfType(0 To 1) As Byte
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'BMP文件头
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'调色板结构
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

'调色板
Private ColorTable() As RGBQUAD

'从二进制数组读取BMP文件
Public Function LoadAnyBitmapFromStream(ByRef Bytes() As Byte, _
                                         ByRef Bitmap() As Byte, _
                                         ByRef Width As Long, _
                                         ByRef Height As Long, _
                                         ByRef BitCount As Long, _
                                         ByRef DibWidth As Long) As Boolean
        Dim Index As Long, Length As Long
        'BMP文件头
        Dim HeaderBitmap As BITMAPFILEHEADER
        '定位内存复制启始索引及长度
        Index = LBound(Bytes)
        Length = Len(HeaderBitmap)
        '复制文件头
        CopyMemoryLong ByVal VarPtr(HeaderBitmap), ByVal VarPtr(Bytes(Index)), Length
        With HeaderBitmap
                '判断是否为BMP文件
                If Chr(.bfType(0)) & Chr(.bfType(1)) <> "BM" Then
                        Err.Description = "非BMP图像"
                        Err.Number = 800
                        Err.Raise 800
                        Exit Function
                End If
        End With
        'BMP信息
        Dim InfoBitmap As BITMAPINFOHEADER
        '定位内存复制启始索引及长度
        Index = Index + Length
        Length = Len(InfoBitmap)
        '复制信息结构
        CopyMemoryLong ByVal VarPtr(InfoBitmap), ByVal VarPtr(Bytes(Index)), Length
        '取BMP信息
        With InfoBitmap
                '判断是否为非压缩24位BMP位图
                If .biCompression <> 0 Then
                        'RLE压缩方式
                        Err.Description = "不支持RLE压缩方式存放的BMP格式"
                        Err.Number = 801
                        Err.Raise 801
                        Exit Function
                End If
                '返回颜色深度
                BitCount = .biBitCount
                '返回BMP尺寸参数
                Width = .biWidth:
                Height = .biHeight
                
                '计算DIB宽度
                Select Case BitCount
                Case 1
                        DibWidth = (((.biWidth + 7) \ 8 + 3) And &HFFFFFFFC)
                Case 4
                        DibWidth = (((.biWidth + 1) \ 2 + 3) And &HFFFFFFFC)
                Case 8
                        DibWidth = ((.biWidth + 3) And &HFFFFFFFC)
                Case 16
                        DibWidth = ((.biWidth * 2 + 3) And &HFFFFFFFC)
                Case 24
                        DibWidth = ((.biWidth * 3 + 3) And &HFFFFFFFC)
                Case 32
                        DibWidth = .biWidth * 4
                Case Else
                        Err.Description = "未知的颜色深度描述"
                        Err.Number = 802
                        Err.Raise 802
                        Exit Function
                End Select
                
                '计算图像大小
                Dim biSize As Long
                biSize = DibWidth * .biHeight
                
                'BMP数据数组分配内存
                ReDim Bitmap(0 To biSize - 1)
                
                '计算偏移值
                Index = Index + Len(InfoBitmap)
                
                '只有biBitCount等于1、4、8时才有调色板。调色板实际上是一个数组，元素的个数由biBitCount和biClrUsed决定。
                '颜色表的大小根据所使用的颜色模式而定：
                '2色图像为8字节；
                '16色图像位64字节；
                '256色图像为1024字节。
                
                '其中，每4字节表示一种颜色，并以B（蓝色）、G（绿色）、R（红色）、alpha（32位位图的透明度值，一般不需要）。
                '即首先4字节表示颜色号0的颜色，接下来表示颜色号1的颜色，依此类推。
                
                '调色板数据
                Select Case BitCount
                Case 1
                        ReDim ColorTable(0 To 1)
                        CopyMemoryLong ByVal VarPtr(ColorTable(0).rgbBlue), ByVal VarPtr(Bytes(Index)), 8
                        Index = Index + 8
                Case 4
                        ReDim ColorTable(0 To 15)
                        CopyMemoryLong ByVal VarPtr(ColorTable(0).rgbBlue), ByVal VarPtr(Bytes(Index)), 64
                        Index = Index + 64
                Case 8
                        ReDim ColorTable(0 To 256)
                        CopyMemoryLong ByVal VarPtr(ColorTable(0).rgbBlue), ByVal VarPtr(Bytes(Index)), 1024
                        Index = Index + 1024
                
                Case 16
                        Const BI_RGB = 0
                        Const BI_bitfields = 3&
                        
                        Select Case .biCompression
                        Case BI_RGB
                                
                                '无调色板
                        Case BI_bitfields
                                '首3个DWORD为 RGB掩码
                                '别用于描述红、绿、蓝分量在16位中所占的位置
                                'RGB 555    0x7C00 0x03E0 0x001F
                                'RGB 565    0xF800 0x07E0 0x001F
                        
                                '你在读取一个像素之后，
                                '可以分别用掩码“与”上像素值，从而提取出想要的颜色分量
                                
                                Err.Description = "不支持16位 BI_BITFIELDS 调色板模式"
                                Err.Number = 803
                                Err.Raise 803
                                Exit Function
                        End Select
                        
                Case 24 '无调色板
                Case 32 '无调色板
                        
                End Select
                
                '计算长度
                Length = biSize
                '复制数据
                CopyMemoryLong ByVal VarPtr(Bitmap(0)), ByVal VarPtr(Bytes(Index)), Length
        End With
        '返回值
        LoadAnyBitmapFromStream = True
End Function


Private Sub Form_Load()
        With ComboBMPS
                .AddItem "1"
                .AddItem "4"
                .AddItem "8"
                .AddItem "16"
                .AddItem "24"
                .AddItem "32"
                .ListIndex = 0
        End With
End Sub

Private Sub ComboBMPS_Click()
        Call PCanvas_DblClick
End Sub


'分析非24BitBMP
Private Sub PCanvas_DblClick()
        PCanvas.Enabled = False
        
        Dim FileName As String, Bytes() As Byte
        
        
        FileName = App.Path & "\..\Example" & ComboBMPS.Text & ".bmp"
        
        
        If Dir(FileName) = vbNullString Then End
        Open FileName For Binary As #1
                ReDim Bytes(LOF(1))
                Get #1, , Bytes
        Close #1
        
        Dim BMP() As Byte, biW As Long, biH As Long, biBitCount As Long, biDIB_W As Long
        '加载任意色深的BMP
        If LoadAnyBitmapFromStream(Bytes, BMP, biW, biH, biBitCount, biDIB_W) = False Then End
        
        
        Dim i As Long, j As Long
        Dim lColor As Long
        Dim R As Byte, G As Byte, B As Byte
        Dim Index As Long
        '使用索引累加法提高坐标转换效率
        Dim X As Long, Y As Long
        X = 0: Y = 0
        
        '是否自动刷新
        PCanvas.AutoRedraw = IIf(CheckDBB.Value = vbChecked, True, False)
        
        Select Case biBitCount
        Case 1 '用1位表示一个像素，所以一个字节可以表示8个像素。坐标是从最左边（最高位）开始的，而不是一般情况下的最低位
                '二进制位表
                Dim BitTable(1 To 8) As Byte
                For i = LBound(BMP) To UBound(BMP)
                        '取一个字节
                        Index = BMP(i)
                        
                        '计算当前字节的各个二进制位
                        For j = 1 To 8
                                If (Index Mod 2) = 1 Then
                                        BitTable(9 - j) = 1
                                Else
                                        BitTable(9 - j) = 0
                                End If
                                Index = Index \ 2
                        Next j
                        
                        For j = 1 To 8
                                '根据二进制位获取颜色
                                With ColorTable(BitTable(j))
                                        B = .rgbBlue
                                        G = .rgbGreen
                                        R = .rgbRed
                                End With
                                SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                                '索引转坐标
                                X = X + 1
                                If X = biW Then
                                        X = 0
                                        Y = Y + 1
                                End If
                        Next j
                Next i
        Case 4 '用4位表示一个像素，所以一个字节可以表示2个像素。坐标是从最左边（最高位）开始的，而不是一般情况下的最低位
        
              'biBitCount=4 表示位图最多有16种颜色。
              '每个象素用4位表示，
              '并用这4位作为彩色表的表项来查找该象素的颜色。
              '例如，如果位图中的第一个字节为0×1F，它表示有
              '两个象素，第一象素的颜色就在彩色表的第2表项中
              '查找，而第二个象素的颜色就在彩色表的第16表项中
              '查找。此时，调色板中缺省情况下会有16个RGB项。对应于索引0到索引15。
                
                '高位在前,低位在后,按索引查表
                For i = LBound(BMP) To UBound(BMP)
                        '取最高4位
                        '[11111111]  \  [00010000]
                        With ColorTable(BMP(i) \ 16)
                                B = .rgbBlue
                                G = .rgbGreen
                                R = .rgbRed
                        End With
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                        
                        '取低4位
                        '[11111111] And [00001111]
                        With ColorTable(BMP(i) And 15)
                                B = .rgbBlue
                                G = .rgbGreen
                                R = .rgbRed
                        End With
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                Next i
        Case 8 '用8位表示一个像素，所以一个字节刚好只能表示一个像素
                For i = LBound(BMP) To UBound(BMP)
                        '获取当前颜色
                        lColor = BMP(i)
                        
                        '是否采用手动调色板
                        If False Then
                                '取RGB分量
                                   '取最高3位
                                   '[11111111]  \  [00100000]
                                B = lColor \ 32
                                   '取低5位
                                   '[11111111] And [00011111]
                                   '取高3位
                                   '[00011111]  \  [00000100]
                                G = (lColor And 31) \ 4
                                   '取低2位
                                   '[11111111] And [00000011]
                                R = lColor And 3
                                '8位RGB332转24位RGB888
                                '111
                                '111
                                '11
                                B = B * 32
                                G = G * 32
                                R = R * 64
                        Else
                        '采用系统调色板
                                With ColorTable(lColor)
                                        R = .rgbRed
                                        G = .rgbGreen
                                        B = .rgbBlue
                                End With
                        End If
                        
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '延时观察
                        If False Then
                                For j = 0 To 20000
                                Next j
                        End If
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                Next i
        Case 16 '用16位表示一个像素，所以两个字节可以表示1个像素。
                '默认情况下16位DIB是555格式，最高位无效（这对VB是个福音，因为VB没有16位无符号型）。
                '在内存的摆放形式如下（PC机是低字节在前）：
                
                
                
                'biBitCount=16 表示位图最多有216种颜色。
                '每个色素用16位（2个字节）表示。
                '这种格式叫作高彩色，或叫增强型16位色，或64K色。
                '它的情况比较复杂，
                
                '当biCompression成员的值是BI_RGB时，它没有调色板。
                        '16位中，最低的5位表示蓝色分量，
                        '中间的5位表示绿色分量，
                        '高的5位表示红色分量，
                        '一共占用了15位，最高的一位保留，设为0。
                        '这种格式也被称作555 16位位图。
                
                '如果biCompression成员的值是BI_BITFIELDS，那么情况就复杂了，
                        '首先是原来调色板的位置被三个DWORD变量占据，
                        '称为红、绿、蓝掩码。分别用于描述红、绿、蓝分量在16位中所占的位置。
                        '在Windows 95（或98）中，
                        '系统可接受两种格式的位域：555和565，
                        '在555格式下，红、绿、蓝的掩码分别是：
                        '0×7C00、0×03E0、0×001F，
                        '而在565格式下，它们则分别为：
                        '0xF800、0×07E0、0×001F。你在读取一个像素之后，
                        '可以分别用掩码“与”上像素值，从而提取出想要的颜色分
                        '量（当然还要再经过适当的左右移操作）。
                        '在NT系统中，则没有格式限制，
                        '只不过要求掩码之间不能有重叠。
                        '（注：这种格式的图像使用起来是比较麻烦的，
                        '不过因为它的显示效果接近于真彩，而图像数据又比真彩图像小的多，所以，它更多的被用于游戏软件）。
                
                
                '上下限
                Const Min = 0 '(只能为0)
                Const Max = 31
                Const MaxOverFlow As Long = Max + 1
                '源数            目标变量
                Dim Src As Long, UChar As Byte
                
                
                For i = LBound(BMP) To UBound(BMP) Step 2
                        '生成16位索引 高字节       低字节
                        lColor = BMP(i + 1) * 256 + BMP(i)
                        
                        '取RGB分量
                           '取最高6位
                           '[1111111111111111]  \  [0000010000000000]
                           '取低5位
                           '[1111110000000000] And [0000000000011111]
                        R = (lColor \ 1024) And 31
                           '取低10位
                           '[1111111111111111] And [0000001111111111]
                           '取高5位
                           '[0000001111111111]  \  [0000000000011111]
                        G = (lColor And 1023) \ 31
                           '取低5位
                           '[1111111111111111] And [0000000000011111]
                        B = lColor And 31
                        
                        
                        Src = B
                        '饱和运算
                        UChar = ((Src And (Src >= Min) Or (Src >= MaxOverFlow)) And Max)
                        B = UChar
                        
                        Src = G
                        '饱和运算
                        UChar = ((Src And (Src >= Min) Or (Src >= MaxOverFlow)) And Max)
                        G = UChar
                        
                        Src = R
                        '饱和运算
                        UChar = ((Src And (Src >= Min) Or (Src >= MaxOverFlow)) And Max)
                        R = UChar
                        
                        '16位RGB555转24位RGB888
                        '11111
                        '11111
                        '11111
                        B = B * 8
                        G = G * 8
                        R = R * 8
                        
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '延时观察
                        If False Then
                                For j = 0 To 20000
                                Next j
                        End If
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                Next i
        Case 24  '用24位表示一个像素，所以三个字节可以表示1个像素
                For i = LBound(BMP) To UBound(BMP) - 2 Step 3
                        '取RGB分量
                        B = BMP(i + 0)
                        G = BMP(i + 1)
                        R = BMP(i + 2)
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '延时观察
                        If False Then
                                For j = 0 To 20000
                                Next j
                        End If
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                Next i
        Case 32 '4字节表示一个像素,RGBA 最高位为A(alpha透明度通道)
                For i = LBound(BMP) To UBound(BMP) - 3 Step 4
                        '取RGB分量
                        B = BMP(i + 0)
                        G = BMP(i + 1)
                        R = BMP(i + 2)
                        SetPixelV PCanvas.hDC, X, biH - Y - 1, RGB(R, G, B)
                        '延时观察
                        If False Then
                                For j = 0 To 20000
                                Next j
                        End If
                        '索引转坐标
                        X = X + 1
                        If X = biW Then
                                X = 0
                                Y = Y + 1
                        End If
                Next i
        End Select
        With PCanvas
                If .AutoRedraw = True Then
                        Set .Picture = .Image
                        .Refresh
                        Clipboard.Clear
                        Clipboard.SetData .Picture, vbCFBitmap
                End If
                .Enabled = True
        End With
End Sub
