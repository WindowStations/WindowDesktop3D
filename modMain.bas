Attribute VB_Name = "modMain"
Option Explicit
Public Const MaxSpeedCrouch             As Single = 0.02
Public Const MaxSpeedUnCrouch           As Single = 0.1
Public Const MaxSpeedFly                As Single = 0.1
Public Const JumpSpeedY                 As Single = 0.075
Public Const MaxForceUnCrouch           As Single = 0.002
Public Const MaxForceFly                As Single = 0.0007
Public Const MaxForceCrouch             As Single = 0.002
Public Const HeadHeight                 As Single = 0.8
Public Const CrouchHeight               As Single = 2
Public Const UnCrouchHeight             As Single = 4
Public Const DS_TRUE                    As Long = 1
Public Const DS_FALSE                   As Long = 0
Public Const Pi                         As Single = 3.141593!
Public Const D3D_TRUE                   As Long = 1
Public Const D3D_FALSE                  As Long = 0
Public Const ERROR_DEVICE_NOT_CONNECTED As Long = 1167
Public Const ERROR_SUCCESS              As Long = 0
Public Const ERROR_EMPTY                As Long = 4306
Public LeftThumbDead               As Boolean ': LeftThumbDead = True
Public RightThumbDead              As Boolean ': RightThumbDead = True
Public oldlv                       As Vector2
Public oldrv                       As Vector2
Public Type Vector2
    x As Double
    y As Double
End Type
Public modifierkeydown             As Boolean ': modifierkeydown = False
Public modifieractive             As Boolean ': modifierkeydown = False
Private Type XINPUT_GAMEPAD
    wButtons As Integer
    bLeftTrigger As Byte
    bRightTrigger As Byte
    sThumbLX As Integer
    sThumbLY As Integer
    sThumbRX As Integer
    sThumbRY As Integer
End Type
Public Type XINPUT_STATE
    PacketNumber As Long
    gamepad As XINPUT_GAMEPAD
End Type
Public Enum CONST_DSBCAPS
    DSBCAPS_PRIMARYBUFFER = &H1
    DSBCAPS_STATIC = &H2
    DSBCAPS_LOCHARDWARE = &H4
    DSBCAPS_LOCSOFTWARE = &H8
    DSBCAPS_CTRL3D = &H10
    DSBCAPS_CTRLFREQUENCY = &H20
    DSBCAPS_CTRLPAN = &H40
    DSBCAPS_CTRLVOLUME = &H80
    DSBCAPS_CTRLPOSITIONNOTIFY = &H100
    DSBCAPS_CTRLFX = &H200
    DSBCAPS_STICKYFOCUS = &H4000
    DSBCAPS_GLOBALFOCUS = &H8000&
    DSBCAPS_GETCURRENTPOSITION2 = &H10000
    DSBCAPS_MUTE3DATMAXDISTANCE = &H20000
    DSBCAPS_LOCDEFER = &H40000
End Enum
Public Enum CONST_DSSCL
    DSSCL_NORMAL = &H1
    DSSCL_PRIORITY = &H2
    DSSCL_EXCLUSIVE = &H3
    DSSCL_WRITEPRIMARY = &H4
End Enum
Public Enum D3DX_FILTER
    D3DX_FILTER_NONE = 1
    D3DX_FILTER_POINT = 2
    D3DX_FILTER_LINEAR = 3
    D3DX_FILTER_TRIANGLE = 4
    D3DX_FILTER_BOX = 5
    D3DX_FILTER_MIRROR_U = &H10000
    D3DX_FILTER_MIRROR_V = &H20000
    D3DX_FILTER_MIRROR_W = &H40000
    D3DX_FILTER_MIRROR = &H70000
    D3DX_FILTER_DITHER = &H80000
    D3DX_FILTER_DITHER_DIFFUSION = &H100000
    D3DX_FILTER_SRGB_IN = &H200000
    D3DX_FILTER_SRGB_OUT = &H400000
    D3DX_FILTER_SRGB = &H600000
End Enum
Public Enum D3DXIMAGE_FILEFORMAT
    D3DXIFF_BMP = 0
    D3DXIFF_JPG = 1
    D3DXIFF_TGA = 2
    D3DXIFF_PNG = 3
    D3DXIFF_DDS = 4
    D3DXIFF_PPM = 5
    D3DXIFF_DIB = 6
    D3DXIFF_HDR = 7
    D3DXIFF_PFM = 8
End Enum
Public Type D3DXIMAGE_INFO
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As D3DFORMAT
    ResourceType As D3DRESOURCETYPE
    ImageFileFormat As D3DXIMAGE_FILEFORMAT
End Type
Public Enum D3DVERTEXTEXTURESAMPLER
    D3DVERTEXTEXTURESAMPLER0 = &H101
    D3DVERTEXTEXTURESAMPLER1 = &H102
    D3DVERTEXTEXTURESAMPLER2 = &H103
    D3DVERTEXTEXTURESAMPLER3 = &H104
End Enum
Public Enum HRESULT
    D3DERR_DEVICELOST = &H88760868
    D3DERR_DEVICENOTRESET = &H88760869
    D3DERR_DRIVERINTERNALERROR = &H88760827
End Enum
Public Enum D3DVERTEXBLENDFLAGS
    D3DVBF_DISABLE = 0
    D3DVBF_1WEIGHTS = 1
    D3DVBF_2WEIGHTS = 2
    D3DVBF_3WEIGHTS = 3
    D3DVBF_TWEENING = 255
    D3DVBF_0WEIGHTS = 256
End Enum
Public Enum D3DCMPFUNC
    D3DCMP_NEVER = 1
    D3DCMP_LESS = 2
    D3DCMP_EQUAL = 3
    D3DCMP_LESSEQUAL = 4
    D3DCMP_GREATER = 5
    D3DCMP_NOTEQUAL = 6
    D3DCMP_GREATEREQUAL = 7
    D3DCMP_ALWAYS = 8
End Enum
Public Enum D3DTEXTURETRANSFORMFLAGS
    D3DTTFF_DISABLE = 0
    D3DTTFF_COUNT1 = 1
    D3DTTFF_COUNT2 = 2
    D3DTTFF_COUNT3 = 3
    D3DTTFF_COUNT4 = 4
    D3DTTFF_PROJECTED = 256
End Enum
Public Enum D3DUSAGE
    D3DUSAGE_NONE = &H0
    D3DUSAGE_RENDERTARGET = &H1
    D3DUSAGE_DEPTHSTENCIL = &H2
    D3DUSAGE_DYNAMIC = &H200
    D3DUSAGE_AUTOGENMIPMAP = &H400
End Enum
Public Enum D3DTEXTUREADDRESS
    D3DTADDRESS_WRAP = 1
    D3DTADDRESS_MIRROR = 2
    D3DTADDRESS_CLAMP = 3
    D3DTADDRESS_BORDER = 4
    D3DTADDRESS_MIRRORONCE = 5
End Enum
Public Enum D3DTEXTUREFILTERTYPE
    D3DTEXF_NONE = 0
    D3DTEXF_POINT = 1
    D3DTEXF_LINEAR = 2
    D3DTEXF_ANISOTROPIC = 3
    D3DTEXF_PYRAMIDALQUAD = 6
    D3DTEXF_GAUSSIANQUAD = 7
    D3DTEXF_CONVOLUTIONMONO = 8
End Enum
Public Enum D3DTA
    D3DTA_SELECTMASK = &HF
    D3DTA_DIFFUSE = &H0
    D3DTA_CURRENT = &H1
    D3DTA_TEXTURE = &H2
    D3DTA_TFACTOR = &H3
    D3DTA_SPECULAR = &H4
    D3DTA_TEMP = &H5
    D3DTA_CONSTANT = &H6
    D3DTA_COMPLEMENT = &H10
    D3DTA_ALPHAREPLICATE = &H20
End Enum
Public Enum D3DTEXTUREOP
    D3DTOP_DISABLE = 1
    D3DTOP_SELECTARG1 = 2
    D3DTOP_SELECTARG2 = 3
    D3DTOP_MODULATE = 4
    D3DTOP_MODULATE2X = 5
    D3DTOP_MODULATE4X = 6
    D3DTOP_ADD = 7
    D3DTOP_ADDSIGNED = 8
    D3DTOP_ADDSIGNED2X = 9
    D3DTOP_SUBTRACT = 10
    D3DTOP_ADDSMOOTH = 11
    D3DTOP_BLENDDIFFUSEALPHA = 12
    D3DTOP_BLENDTEXTUREALPHA = 13
    D3DTOP_BLENDFACTORALPHA = 14
    D3DTOP_BLENDTEXTUREALPHAPM = 15
    D3DTOP_BLENDCURRENTALPHA = 16
    D3DTOP_PREMODULATE = 17
    D3DTOP_MODULATEALPHA_ADDCOLOR = 18
    D3DTOP_MODULATECOLOR_ADDALPHA = 19
    D3DTOP_MODULATEINVALPHA_ADDCOLOR = 20
    D3DTOP_MODULATEINVCOLOR_ADDALPHA = 21
    D3DTOP_BUMPENVMAP = 22
    D3DTOP_BUMPENVMAPLUMINANCE = 23
    D3DTOP_DOTPRODUCT3 = 24
    D3DTOP_MULTIPLYADD = 25
    D3DTOP_LERP = 26
End Enum
Public Enum D3DBLENDOP
    D3DBLENDOP_ADD = 1
    D3DBLENDOP_SUBTRACT = 2
    D3DBLENDOP_REVSUBTRACT = 3
    D3DBLENDOP_MIN = 4
    D3DBLENDOP_MAX = 5
End Enum
Public Enum D3DBLEND
    D3DBLEND_ZERO = 1
    D3DBLEND_ONE = 2
    D3DBLEND_SRCCOLOR = 3
    D3DBLEND_INVSRCCOLOR = 4
    D3DBLEND_SRCALPHA = 5
    D3DBLEND_INVSRCALPHA = 6
    D3DBLEND_DESTALPHA = 7
    D3DBLEND_INVDESTALPHA = 8
    D3DBLEND_DESTCOLOR = 9
    D3DBLEND_INVDESTCOLOR = 10
    D3DBLEND_SRCALPHASAT = 11
    D3DBLEND_BOTHSRCALPHA = 12
    D3DBLEND_BOTHINVSRCALPHA = 13
    D3DBLEND_BLENDFACTOR = 14
    D3DBLEND_INVBLENDFACTOR = 15
    D3DBLEND_SRCCOLOR2 = 16
    D3DBLEND_INVSRCCOLOR2 = 17
End Enum
Public Enum D3DTEXTURESTAGESTATETYPE
    D3DTSS_COLOROP = 1
    D3DTSS_COLORARG1 = 2
    D3DTSS_COLORARG2 = 3
    D3DTSS_ALPHAOP = 4
    D3DTSS_ALPHAARG1 = 5
    D3DTSS_ALPHAARG2 = 6
    D3DTSS_BUMPENVMAT00 = 7
    D3DTSS_BUMPENVMAT01 = 8
    D3DTSS_BUMPENVMAT10 = 9
    D3DTSS_BUMPENVMAT11 = 10
    D3DTSS_TEXCOORDINDEX = 11
    D3DTSS_BUMPENVLSCALE = 22
    D3DTSS_BUMPENVLOFFSET = 23
    D3DTSS_TEXTURETRANSFORMFLAGS = 24
    D3DTSS_COLORARG0 = 26
    D3DTSS_ALPHAARG0 = 27
    D3DTSS_RESULTARG = 28
    D3DTSS_CONSTANT = 32
End Enum
Public Enum D3DSAMPLERSTATETYPE
    D3DSAMP_ADDRESSU = 1
    D3DSAMP_ADDRESSV = 2
    D3DSAMP_ADDRESSW = 3
    D3DSAMP_BORDERCOLOR = 4
    D3DSAMP_MAGFILTER = 5
    D3DSAMP_MINFILTER = 6
    D3DSAMP_MIPFILTER = 7
    D3DSAMP_MIPMAPLODBIAS = 8
    D3DSAMP_MAXMIPLEVEL = 9
    D3DSAMP_MAXANISOTROPY = 10
    D3DSAMP_SRGBTEXTURE = 11
    D3DSAMP_ELEMENTINDEX = 12
    D3DSAMP_DMAPOFFSET = 13
End Enum
Public Enum D3DTRANSFORMSTATETYPE
    D3DTS_VIEW = 2
    D3DTS_PROJECTION = 3
    D3DTS_TEXTURE0 = 16
    D3DTS_TEXTURE1 = 17
    D3DTS_TEXTURE2 = 18
    D3DTS_TEXTURE3 = 19
    D3DTS_TEXTURE4 = 20
    D3DTS_TEXTURE5 = 21
    D3DTS_TEXTURE6 = 22
    D3DTS_TEXTURE7 = 23
    D3DTS_WORLD = 256
End Enum
Public Enum D3DFORMAT
    D3DFMT_UNKNOWN = 0
    D3DFMT_R8G8B8 = 20
    D3DFMT_A8R8G8B8 = 21
    D3DFMT_X8R8G8B8 = 22
    D3DFMT_R5G6B5 = 23
    D3DFMT_X1R5G5B5 = 24
    D3DFMT_A1R5G5B5 = 25
    D3DFMT_A4R4G4B4 = 26
    D3DFMT_R3G3B2 = 27
    D3DFMT_A8 = 28
    D3DFMT_A8R3G3B2 = 29
    D3DFMT_X4R4G4B4 = 30
    D3DFMT_A2B10G10R10 = 31
    D3DFMT_A8B8G8R8 = 32
    D3DFMT_X8B8G8R8 = 33
    D3DFMT_G16R16 = 34
    D3DFMT_A2R10G10B10 = 35
    D3DFMT_A16B16G16R16 = 36
    D3DFMT_A8P8 = 40
    D3DFMT_P8 = 41
    D3DFMT_L8 = 50
    D3DFMT_A8L8 = 51
    D3DFMT_A4L4 = 52
    D3DFMT_V8U8 = 60
    D3DFMT_L6V5U5 = 61
    D3DFMT_X8L8V8U8 = 62
    D3DFMT_Q8W8V8U8 = 63
    D3DFMT_V16U16 = 64
    D3DFMT_A2W10V10U10 = 67
    's="UYVY":?"&h";right$("0" & hex$(asc(mid(s,4,1))),2);right$("0" & hex$(asc(mid(s,3,1))),2);right$("0" & hex$(asc(mid(s,2,1))),2);right$("0" & hex$(asc(mid(s,1,1))),2);"&"
    D3DFMT_UYVY = &H59565955 '               = MAKEFOURCC('U', 'Y', 'V', 'Y'),
    D3DFMT_R8G8_B8G8 = &H47424752  '         = MAKEFOURCC('R', 'G', 'B', 'G'),
    D3DFMT_YUY2 = &H32595559 '               = MAKEFOURCC('Y', 'U', 'Y', '2'),
    D3DFMT_G8R8_G8B8 = &H42475247 '          = MAKEFOURCC('G', 'R', 'G', 'B'),
    D3DFMT_DXT1 = &H31545844 '               = MAKEFOURCC('D', 'X', 'T', '1'),
    D3DFMT_DXT2 = &H32545844 '               = MAKEFOURCC('D', 'X', 'T', '2'),
    D3DFMT_DXT3 = &H33545844 '               = MAKEFOURCC('D', 'X', 'T', '3'),
    D3DFMT_DXT4 = &H34545844 '               = MAKEFOURCC('D', 'X', 'T', '4'),
    D3DFMT_DXT5 = &H35545844 '               = MAKEFOURCC('D', 'X', 'T', '5'),
    D3DFMT_D16_LOCKABLE = 70
    D3DFMT_D32 = 71
    D3DFMT_D15S1 = 73
    D3DFMT_D24S8 = 75
    D3DFMT_D24X8 = 77
    D3DFMT_D24X4S4 = 79
    D3DFMT_D16 = 80
    D3DFMT_D32F_LOCKABLE = 82
    D3DFMT_D24FS8 = 83
    D3DFMT_D32_LOCKABLE = 84
    D3DFMT_S8_LOCKABLE = 85
    D3DFMT_L16 = 81
    D3DFMT_VERTEXDATA = 100
    D3DFMT_INDEX16 = 101
    D3DFMT_INDEX32 = 102
    D3DFMT_Q16W16V16U16 = 110
    D3DFMT_MULTI2_ARGB8 = &H3154454D '       = MAKEFOURCC('M','E','T','1'),
    D3DFMT_R16F = 111
    D3DFMT_G16R16F = 112
    D3DFMT_A16B16G16R16F = 113
    D3DFMT_R32F = 114
    D3DFMT_G32R32F = 115
    D3DFMT_A32B32G32R32F = 116
    D3DFMT_CxV8U8 = 117
    D3DFMT_A1 = 118
    D3DFMT_BINARYBUFFER = 199
End Enum
Public Enum D3DMULTISAMPLE_TYPE
    D3DMULTISAMPLE_NONE = 0
    D3DMULTISAMPLE_NONMASKABLE = 1
    D3DMULTISAMPLE_2_SAMPLES = 2
    D3DMULTISAMPLE_3_SAMPLES = 3
    D3DMULTISAMPLE_4_SAMPLES = 4
    D3DMULTISAMPLE_5_SAMPLES = 5
    D3DMULTISAMPLE_6_SAMPLES = 6
    D3DMULTISAMPLE_7_SAMPLES = 7
    D3DMULTISAMPLE_8_SAMPLES = 8
    D3DMULTISAMPLE_9__SAMPLES = 9
    D3DMULTISAMPLE_10_SAMPLES = 10
    D3DMULTISAMPLE_11_SAMPLES = 11
    D3DMULTISAMPLE_12_SAMPLES = 12
    D3DMULTISAMPLE_13_SAMPLES = 13
    D3DMULTISAMPLE_14_SAMPLES = 14
    D3DMULTISAMPLE_15_SAMPLES = 15
    D3DMULTISAMPLE_16_SAMPLES = 16
End Enum
Public Enum D3DSWAPEFFECT
    D3DSWAPEFFECT_DISCARD = 1
    D3DSWAPEFFECT_FLIP = 2
    D3DSWAPEFFECT_COPY = 3
End Enum
Public Enum CONST_D3DCREATEFLAGS
    D3DCREATE_FPU_PRESERVE = 2
    D3DCREATE_MULTITHREADED = 4
    D3DCREATE_PUREDEVICE = 16
    D3DCREATE_SOFTWARE_VERTEXPROCESSING = 32
    D3DCREATE_HARDWARE_VERTEXPROCESSING = 64
    D3DCREATE_MIXED_VERTEXPROCESSING = 128
    D3DCREATE_DISABLE_DRIVER_MANAGEMENT = &H100
    D3DCREATE_ADAPTERGROUP_DEVICE = &H200
    D3DCREATE_DISABLE_DRIVER_MANAGEMENT_EX = &H400
    D3DCREATE_NOWINDOWCHANGES = &H800
    D3DCREATE_DISABLE_PSGP_THREADING = &H2000
    D3DCREATE_ENABLE_PRESENTSTATS = &H4000
    D3DCREATE_DISABLE_PRINTSCREEN = &H8000&
    D3DCREATE_SCREENSAVER = &H10000000
End Enum
Public Enum D3DDEVTYPE
    D3DDEVTYPE_HAL = 1
    D3DDEVTYPE_REF = 2
    D3DDEVTYPE_SW = 3
    D3DDEVTYPE_NULLREF = 4
End Enum
Public Enum PRESENTATION_INTERVAL
    D3DPRESENT_DEFAULT = &H0
    D3DPRESENT_ONE = &H1
    D3DPRESENT_TWO = &H2
    D3DPRESENT_THREE = &H4
    D3DPRESENT_FOUR = &H8
    D3DPRESENT_IMMEDIATE = &H80000000
End Enum
Public Enum CONST_D3DCLEARFLAGS
    D3DCLEAR_TARGET = &H1
    D3DCLEAR_ZBUFFER = &H2
    D3DCLEAR_STENCIL = &H4
End Enum
Public Enum CONST_D3DFVF
    D3DFVF_RESERVED0 = &H1
    D3DFVF_POSITION_MASK = &H400E
    D3DFVF_XYZ = &H2
    D3DFVF_XYZRHW = &H4
    D3DFVF_XYZB1 = &H6
    D3DFVF_XYZB2 = &H8
    D3DFVF_XYZB3 = &HA
    D3DFVF_XYZB4 = &HC
    D3DFVF_XYZB5 = &HE
    D3DFVF_XYZW = &H4002
    D3DFVF_NORMAL = &H10
    D3DFVF_PSIZE = &H20
    D3DFVF_DIFFUSE = &H40
    D3DFVF_SPECULAR = &H80
    D3DFVF_TEXCOUNT_MASK = &HF00
    D3DFVF_TEXCOUNT_SHIFT = 8
    D3DFVF_TEX0 = &H0
    D3DFVF_TEX1 = &H100
    D3DFVF_TEX2 = &H200
    D3DFVF_TEX3 = &H300
    D3DFVF_TEX4 = &H400
    D3DFVF_TEX5 = &H500
    D3DFVF_TEX6 = &H600
    D3DFVF_TEX7 = &H700
    D3DFVF_TEX8 = &H800
    D3DFVF_LASTBETA_UBYTE4 = &H1000
    D3DFVF_LASTBETA_D3DCOLORVALUE = &H8000&
    D3DFVF_RESERVED2 = &H6000
End Enum
Public Enum D3DPOOL
    D3DPOOL_DEFAULT = 0
    D3DPOOL_MANAGED = 1
    D3DPOOL_SYSTEMMEM = 2
    D3DPOOL_SCRATCH = 3
End Enum
Public Enum D3DPRIMITIVETYPE
    D3DPT_POINTLIST = 1
    D3DPT_LINELIST = 2
    D3DPT_LINESTRIP = 3
    D3DPT_TRIANGLELIST = 4
    D3DPT_TRIANGLESTRIP = 5
    D3DPT_TRIANGLEFAN = 6
End Enum
Public Enum D3DRENDERSTATETYPE
    D3DRS_ZENABLE = 7
    D3DRS_FILLMODE = 8
    D3DRS_SHADEMODE = 9
    D3DRS_ZWRITEENABLE = 14
    D3DRS_ALPHATESTENABLE = 15
    D3DRS_LASTPIXEL = 16
    D3DRS_SRCBLEND = 19
    D3DRS_DESTBLEND = 20
    D3DRS_CULLMODE = 22
    D3DRS_ZFUNC = 23
    D3DRS_ALPHAREF = 24
    D3DRS_ALPHAFUNC = 25
    D3DRS_DITHERENABLE = 26
    D3DRS_ALPHABLENDENABLE = 27
    D3DRS_FOGENABLE = 28
    D3DRS_SPECULARENABLE = 29
    D3DRS_FOGCOLOR = 34
    D3DRS_FOGTABLEMODE = 35
    D3DRS_FOGSTART = 36
    D3DRS_FOGEND = 37
    D3DRS_FOGDENSITY = 38
    D3DRS_RANGEFOGENABLE = 48
    D3DRS_STENCILENABLE = 52
    D3DRS_STENCILFAIL = 53
    D3DRS_STENCILZFAIL = 54
    D3DRS_STENCILPASS = 55
    D3DRS_STENCILFUNC = 56
    D3DRS_STENCILREF = 57
    D3DRS_STENCILMASK = 58
    D3DRS_STENCILWRITEMASK = 59
    D3DRS_TEXTUREFACTOR = 60
    D3DRS_WRAP0 = 128
    D3DRS_WRAP1 = 129
    D3DRS_WRAP2 = 130
    D3DRS_WRAP3 = 131
    D3DRS_WRAP4 = 132
    D3DRS_WRAP5 = 133
    D3DRS_WRAP6 = 134
    D3DRS_WRAP7 = 135
    D3DRS_CLIPPING = 136
    D3DRS_LIGHTING = 137
    D3DRS_AMBIENT = 139
    D3DRS_FOGVERTEXMODE = 140
    D3DRS_COLORVERTEX = 141
    D3DRS_LOCALVIEWER = 142
    D3DRS_NORMALIZENORMALS = 143
    D3DRS_DIFFUSEMATERIALSOURCE = 145
    D3DRS_SPECULARMATERIALSOURCE = 146
    D3DRS_AMBIENTMATERIALSOURCE = 147
    D3DRS_EMISSIVEMATERIALSOURCE = 148
    D3DRS_VERTEXBLEND = 151
    D3DRS_CLIPPLANEENABLE = 152
    D3DRS_POINTSIZE = 154
    D3DRS_POINTSIZE_MIN = 155
    D3DRS_POINTSPRITEENABLE = 156
    D3DRS_POINTSCALEENABLE = 157
    D3DRS_POINTSCALE_A = 158
    D3DRS_POINTSCALE_B = 159
    D3DRS_POINTSCALE_C = 160
    D3DRS_MULTISAMPLEANTIALIAS = 161
    D3DRS_MULTISAMPLEMASK = 162
    D3DRS_PATCHEDGESTYLE = 163
    D3DRS_DEBUGMONITORTOKEN = 165
    D3DRS_POINTSIZE_MAX = 166
    D3DRS_INDEXEDVERTEXBLENDENABLE = 167
    D3DRS_COLORWRITEENABLE = 168
    D3DRS_TWEENFACTOR = 170
    D3DRS_BLENDOP = 171
    D3DRS_POSITIONDEGREE = 172
    D3DRS_NORMALDEGREE = 173
    D3DRS_SCISSORTESTENABLE = 174
    D3DRS_SLOPESCALEDEPTHBIAS = 175
    D3DRS_ANTIALIASEDLINEENABLE = 176
    D3DRS_MINTESSELLATIONLEVEL = 178
    D3DRS_MAXTESSELLATIONLEVEL = 179
    D3DRS_ADAPTIVETESS_X = 180
    D3DRS_ADAPTIVETESS_Y = 181
    D3DRS_ADAPTIVETESS_Z = 182
    D3DRS_ADAPTIVETESS_W = 183
    D3DRS_ENABLEADAPTIVETESSELLATION = 184
    D3DRS_TWOSIDEDSTENCILMODE = 185
    D3DRS_CCW_STENCILFAIL = 186
    D3DRS_CCW_STENCILZFAIL = 187
    D3DRS_CCW_STENCILPASS = 188
    D3DRS_CCW_STENCILFUNC = 189
    D3DRS_COLORWRITEENABLE1 = 190
    D3DRS_COLORWRITEENABLE2 = 191
    D3DRS_COLORWRITEENABLE3 = 192
    D3DRS_BLENDFACTOR = 193
    D3DRS_SRGBWRITEENABLE = 194
    D3DRS_DEPTHBIAS = 195
    D3DRS_WRAP8 = 198
    D3DRS_WRAP9 = 199
    D3DRS_WRAP10 = 200
    D3DRS_WRAP11 = 201
    D3DRS_WRAP12 = 202
    D3DRS_WRAP13 = 203
    D3DRS_WRAP14 = 204
    D3DRS_WRAP15 = 205
    D3DRS_SEPARATEALPHABLENDENABLE = 206
    D3DRS_SRCBLENDALPHA = 207
    D3DRS_DESTBLENDALPHA = 208
    D3DRS_BLENDOPALPHA = 209
End Enum
Public Enum D3DZBUFFERTYPE
    D3DZB_FALSE = 0
    D3DZB_TRUE = 1
    D3DZB_USEW = 2
End Enum
Public Enum D3DFILLMODE
    D3DFILL_POINT = 1
    D3DFILL_WIREFRAME = 2
    D3DFILL_SOLID = 3
End Enum
Public Enum D3DCULL
    D3DCULL_NONE = 1
    D3DCULL_CW = 2
    D3DCULL_CCW = 3
End Enum
Public Enum D3DLIGHTTYPE
    D3DLIGHT_POINT = 1
    D3DLIGHT_SPOT = 2
    D3DLIGHT_DIRECTIONAL = 3
End Enum
Public Enum D3DDECLTYPE
    D3DDECLTYPE_FLOAT1 = 0
    D3DDECLTYPE_FLOAT2 = 1
    D3DDECLTYPE_FLOAT3 = 2
    D3DDECLTYPE_FLOAT4 = 3
    D3DDECLTYPE_D3DCOLORVALUE = 4
    D3DDECLTYPE_UBYTE4 = 5
    D3DDECLTYPE_SHORT2 = 6
    D3DDECLTYPE_SHORT4 = 7
    D3DDECLTYPE_UBYTE4N = 8
    D3DDECLTYPE_SHORT2N = 9
    D3DDECLTYPE_SHORT4N = 10
    D3DDECLTYPE_USHORT2N = 11
    D3DDECLTYPE_USHORT4N = 12
    D3DDECLTYPE_UDEC3 = 13
    D3DDECLTYPE_DEC3N = 14
    D3DDECLTYPE_FLOAT16_2 = 15
    D3DDECLTYPE_FLOAT16_4 = 16
    D3DDECLTYPE_UNUSED = 17
End Enum
Public Enum D3DDECLMETHOD
    D3DDECLMETHOD_DEFAULT = 0
    D3DDECLMETHOD_PARTIALU = 1
    D3DDECLMETHOD_PARTIALV = 2
    D3DDECLMETHOD_CROSSUV = 3
    D3DDECLMETHOD_UV = 4
    D3DDECLMETHOD_LOOKUP = 5
    D3DDECLMETHOD_LOOKUPPRESAMPLED = 6
End Enum
Public Enum D3DDECLUSAGE
    D3DDECLUSAGE_POSITION = 0
    D3DDECLUSAGE_BLENDWEIGHT = 1
    D3DDECLUSAGE_BLENDINDICES = 2
    D3DDECLUSAGE_NORMAL = 3
    D3DDECLUSAGE_PSIZE = 4
    D3DDECLUSAGE_TEXCOORD = 5
    D3DDECLUSAGE_TANGENT = 6
    D3DDECLUSAGE_BINORMAL = 7
    D3DDECLUSAGE_TESSFACTOR = 8
    D3DDECLUSAGE_POSITIONT = 9
    D3DDECLUSAGE_COLOR = 10
    D3DDECLUSAGE_FOG = 11
    D3DDECLUSAGE_DEPTH = 12
    D3DDECLUSAGE_SAMPLE = 13
End Enum
Public Enum D3DRESOURCETYPE
    D3DRTYPE_SURFACE = 1
    D3DRTYPE_VOLUME = 2
    D3DRTYPE_TEXTURE = 3
    D3DRTYPE_VOLUMETEXTURE = 4
    D3DRTYPE_CubeTexture = 5
    D3DRTYPE_VERTEXBUFFER = 6
    D3DRTYPE_INDEXBUFFER = 7
End Enum
Public Enum D3DPTFILTERCAPS
    D3DPTFILTERCAPS_MINFPOINT = &H100&
    D3DPTFILTERCAPS_MINFLINEAR = &H200&
    D3DPTFILTERCAPS_MINFANISOTROPIC = &H400&
    D3DPTFILTERCAPS_MINFPYRAMIDALQUAD = &H800&
    D3DPTFILTERCAPS_MINFGAUSSIANQUAD = &H1000&
    D3DPTFILTERCAPS_MIPFPOINT = &H10000
    D3DPTFILTERCAPS_MIPFLINEAR = &H20000
    D3DPTFILTERCAPS_MAGFPOINT = &H1000000
    D3DPTFILTERCAPS_MAGFLINEAR = &H2000000
    D3DPTFILTERCAPS_MAGFANISOTROPIC = &H4000000
    D3DPTFILTERCAPS_MAGFPYRAMIDALQUAD = &H8000000
    D3DPTFILTERCAPS_MAGFGAUSSIANQUAD = &H10000000
End Enum
Public Enum D3DPTEXTURECAPS
    D3DPTEXTURECAPS_PERSPECTIVE = &H1&
    D3DPTEXTURECAPS_POW2 = &H2&
    D3DPTEXTURECAPS_ALPHA = &H4&
    D3DPTEXTURECAPS_SQUAREONLY = &H20&
    D3DPTEXTURECAPS_TEXREPEATNOTSCALEDBYSIZE = &H40&
    D3DPTEXTURECAPS_ALPHAPALETTE = &H80&
    D3DPTEXTURECAPS_NONPOW2CONDITIONAL = &H100&
    D3DPTEXTURECAPS_PROJECTED = &H400&
    D3DPTEXTURECAPS_CUBEMAP = &H800&
    D3DPTEXTURECAPS_VOLUMEMAP = &H2000&
    D3DPTEXTURECAPS_MIPMAP = &H4000&
    D3DPTEXTURECAPS_MIPVOLUMEMAP = &H8000&
    D3DPTEXTURECAPS_MIPCUBEMAP = &H10000
    D3DPTEXTURECAPS_CUBEMAP_POW2 = &H20000
    D3DPTEXTURECAPS_VOLUMEMAP_POW2 = &H40000
    D3DPTEXTURECAPS_NOPROJECTEDBUMPENV = &H200000
End Enum
Public Type D3DPRESENT_PARAMETERS
    BackBufferWidth As Long
    BackBufferHeight As Long
    BackBufferFormat As D3DFORMAT
    BackBufferCount As Long
    MultiSampleType As D3DMULTISAMPLE_TYPE
    MultiSampleQuality As Long
    SwapEffect As D3DSWAPEFFECT
    hDeviceWindow As Long
    Windowed As Long
    EnableAutoDepthStencil As Long
    AutoDepthStencilFormat As D3DFORMAT
    flags As Long
    FullScreen_RefreshRateInHz As Long
    PresentationInterval As PRESENTATION_INTERVAL
End Type
Public Type D3DMATRIX
    m11 As Single: m12 As Single: m13 As Single: m14 As Single
    m21 As Single: m22 As Single: m23 As Single: m24 As Single
    m31 As Single: m32 As Single: m33 As Single: m34 As Single
    m41 As Single: m42 As Single: m43 As Single: m44 As Single
End Type
Public Type D3DVECTOR2
    x As Single
    y As Single
End Type
Public Type D3DVECTOR
    x As Single
    y As Single
    z As Single
End Type
Public Type D3DVECTOR4
    x As Single
    y As Single
    z As Single
    w As Single
End Type
Public Type D3DPLANE
    a As Single
    b As Single
    c As Single
    d As Single
End Type
Public Type D3DQUATERNION
    x As Single
    y As Single
    z As Single
    w As Single
End Type
Public Type D3DCOLORVALUE
    r As Single
    G As Single
    b As Single
    a As Single
End Type
Type D3DVIEWPORT9
    x As Long
    y As Long
    Width As Long
    Height As Long
    MinZ As Single
    MaxZ As Single
End Type
Public Type D3DMATERIAL9
    Diffuse As D3DCOLORVALUE
    Ambient As D3DCOLORVALUE
    Specular As D3DCOLORVALUE
    Emissive As D3DCOLORVALUE
    Power As Single
End Type
Public Type D3DLIGHT9
    Type As D3DLIGHTTYPE
    Diffuse As D3DCOLORVALUE
    Specular As D3DCOLORVALUE
    Ambient As D3DCOLORVALUE
    Position As D3DVECTOR
    Direction As D3DVECTOR
    Range As Single
    Falloff As Single
    Attenuation0 As Single
    Attenuation1 As Single
    Attenuation2 As Single
    Theta As Single
    Phi As Single
End Type
Public Type D3DVERTEXELEMENT9
    Stream As Integer
    Offset As Integer
    dType As Byte
    Method As Byte
    Usage As Byte
    UsageIndex As Byte
End Type
Public Type D3DRECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public Type D3DBOX
    left As Long
    top As Long
    right As Long
    bottom As Long
    Front As Long
    Back As Long
End Type
Public Type D3DSURFACE_DESC
    Format As D3DFORMAT
    rType As D3DRESOURCETYPE
    Usage As D3DUSAGE
    Pool As D3DPOOL
    MultiSampleType As D3DMULTISAMPLE_TYPE
    MultiSampleQuality As Long
    Width As Long
    Height As Long
End Type
Public Type D3DVSHADERCAPS2_0
    Caps As Long
    DynamicFlowControlDepth As Long
    NumTemps As Long
    StaticFlowControlDepth As Long
End Type
Public Type D3DPSHADERCAPS2_0
    Caps As Long
    DynamicFlowControlDepth As Long
    NumTemps As Long
    StaticFlowControlDepth As Long
    NumInstructionSlots As Long
End Type
Public Type D3DCAPS9
    DeviceType As D3DDEVTYPE
    AdapterOrdinal As Long
    Caps As Long
    Caps2 As Long
    Caps3 As Long
    PresentationIntervals As Long
    CursorCaps As Long
    DevCaps As Long
    PrimitiveMiscCaps As Long
    RASTERCAPS As Long
    ZCmpCaps As Long
    SrcBlendCaps As Long
    DestBlendCaps As Long
    AlphaCmpCaps As Long
    ShadeCaps As Long
    TextureCaps As D3DPTEXTURECAPS
    TextureFilterCaps As D3DPTFILTERCAPS
    CubeTextureFilterCaps As D3DPTFILTERCAPS
    VolumeTextureFilterCaps As D3DPTFILTERCAPS
    TextureAddressCaps As Long
    VolumeTextureAddressCaps As Long
    LineCaps As Long
    MaxTextureWidth As Long
    MaxTextureHeight As Long
    MaxVolumeExtent As Long
    MaxTextureRepeat As Long
    MaxTextureAspectRatio As Long
    MaxAnisotropy As Long
    MaxVertexW As Single
    GuardBandLeft As Single
    GuardBandTop As Single
    GuardBandRight As Single
    GuardBandBottom As Single
    ExtentsAdjust As Single
    StencilCaps As Long
    FVFCaps As Long
    TextureOpCaps As Long
    MaxTextureBlendStages As Long
    MaxSimultaneousTextures As Long
    VertexProcessingCaps As Long
    MaxActiveLights As Long
    MaxUserClipPlanes As Long
    MaxVertexBlendMatrices As Long
    MaxVertexBlendMatrixIndex As Long
    MaxPointSize As Single
    MaxPrimitiveCount As Long
    MaxVertexIndex As Long
    MaxStreams As Long
    MaxStreamStride As Long
    VertexShaderVersion As Long
    MaxVertexShaderConst As Long
    PixelShaderVersion As Long
    PixelShader1xMaxValue As Single
    DevCaps2 As Long
    MaxNpatchTessellationLevel As Single
    Reserved5 As Long
    MasterAdapterOrdinal As Long
    AdapterOrdinalInGroup As Long
    NumberOfAdaptersInGroup As Long
    DeclTypes As Long
    NumSimultaneousRTs As Long
    StretchRectFilterCaps As Long
    VS20Caps As D3DVSHADERCAPS2_0
    PS20Caps As D3DPSHADERCAPS2_0
    VertexTextureFilterCaps As Long
    MaxVShaderInstructionsExecuted As Long
    MaxPShaderInstructionsExecuted As Long
    MaxVertexShader30InstructionSlots As Long
    MaxPixelShader30InstructionSlots As Long
End Type
Private Declare Sub ds_Create Lib "dx9_vb" (ByRef pDS As Long)
Private Declare Sub D3DX_CreateTextureFromFile Lib "dx9_vb" (ByVal pDev As Long, ByVal pSrcFile As Long, ByRef pTex As Long)
Private Declare Sub D3DX_CreateTextureFromFileEx Lib "dx9_vb" (ByVal pDev As Long, ByVal pSrcFile As Long, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As Long, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As Long, ByVal MipFilter As Long, ByVal ColorKey As Long, ByVal pSrcInfo As Long, ByVal pPalette As Long, ByRef pTex As Long)
Private Declare Sub D3DX_GetImageInfoFromFile Lib "dx9_vb" (ByVal pSrcFile As Long, ByRef imgInfo As D3DXIMAGE_INFO)
Public Declare Function Vec2Length Lib "dx9_vb" (ByRef pV As D3DVECTOR2) As Single
Public Declare Function Vec2LengthSq Lib "dx9_vb" (ByRef pV As D3DVECTOR2) As Single
Public Declare Function Vec2Dot Lib "dx9_vb" (ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2) As Single
Public Declare Function Vec2CCW Lib "dx9_vb" (ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2) As Single
Public Declare Sub Vec2Add Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2)
Public Declare Sub Vec2Subtract Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2)
Public Declare Sub Vec2Minimize Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2)
Public Declare Sub Vec2Maximize Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2)
Public Declare Sub Vec2Scale Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV As D3DVECTOR2, ByVal s As Single)
Public Declare Sub Vec2Lerp Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2, ByVal s As Single)
Public Declare Sub Vec2Normalize Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV As D3DVECTOR2)
Public Declare Sub Vec2Hermite Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pT1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2, ByRef pT2 As D3DVECTOR2, ByVal s As Single)
Public Declare Sub Vec2CatmullRom Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV0 As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2, ByRef pV3 As D3DVECTOR2, ByVal s As Single)
Public Declare Sub Vec2BaryCentric Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV1 As D3DVECTOR2, ByRef pV2 As D3DVECTOR2, ByRef pV3 As D3DVECTOR2, ByVal f As Single, ByVal G As Single)
Public Declare Sub Vec2Transform Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV As D3DVECTOR2, ByRef pM As D3DMATRIX)
Public Declare Sub Vec2TransformCoord Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV As D3DVECTOR2, ByRef pM As D3DMATRIX)
Public Declare Sub Vec2TransformNormal Lib "dx9_vb" (ByRef pOut As D3DVECTOR2, ByRef pV As D3DVECTOR2, ByRef pM As D3DMATRIX)
Public Declare Function Vec3Length Lib "dx9_vb" (ByRef pV As D3DVECTOR) As Single
Public Declare Function Vec3LengthSq Lib "dx9_vb" (ByRef pV As D3DVECTOR) As Single
Public Declare Function Vec3Dot Lib "dx9_vb" (ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR) As Single
Public Declare Sub Vec3Cross Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub Vec3Add Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub Vec3Subtract Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub Vec3Minimize Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub Vec3Maximize Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub Vec3Scale Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR, ByVal s As Single)
Public Declare Sub Vec3Lerp Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR, ByVal s As Single)
Public Declare Sub Vec3Normalize Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR)
Public Declare Sub Vec3Hermite Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pT1 As D3DVECTOR, ByRef pV2 As D3DVECTOR, ByRef pT2 As D3DVECTOR, ByVal s As Single)
Public Declare Sub Vec3CatmullRom Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV0 As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR, ByRef pV3 As D3DVECTOR, ByVal s As Single)
Public Declare Sub Vec3BaryCentric Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR, ByRef pV3 As D3DVECTOR, ByVal f As Single, ByVal G As Single)
Public Declare Sub Vec3Transform Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV As D3DVECTOR, ByRef pM As D3DMATRIX)
Public Declare Sub Vec3TransformCoord Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR, ByRef pM As D3DMATRIX)
Public Declare Sub Vec3TransformNormal Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR, ByRef pM As D3DMATRIX)
Public Declare Sub Vec3Project Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR, ByRef pViewport As D3DVIEWPORT9, ByRef pProjection As D3DMATRIX, ByRef pView As D3DMATRIX, ByRef pWorld As D3DMATRIX)
Public Declare Sub Vec3Unproject Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef pV As D3DVECTOR, ByRef pViewport As D3DVIEWPORT9, ByRef pProjection As D3DMATRIX, ByRef pView As D3DMATRIX, ByRef pWorld As D3DMATRIX)
Public Declare Function Vec4Length Lib "dx9_vb" (ByRef pV As D3DVECTOR4) As Single
Public Declare Function Vec4LengthSq Lib "dx9_vb" (ByRef pV As D3DVECTOR4) As Single
Public Declare Function Vec4Dot Lib "dx9_vb" (ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4) As Single
Public Declare Sub Vec4Add Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4)
Public Declare Sub Vec4Subtract Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4)
Public Declare Sub Vec4Minimize Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4)
Public Declare Sub Vec4Maximize Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4)
Public Declare Sub Vec4Scale Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV As D3DVECTOR4, ByVal s As Single)
Public Declare Sub Vec4Lerp Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4, ByVal s As Single)
Public Declare Sub Vec4Cross Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4, ByRef pV3 As D3DVECTOR4)
Public Declare Sub Vec4Normalize Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV As D3DVECTOR4)
Public Declare Sub Vec4Hermite Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pT1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4, ByRef pT2 As D3DVECTOR4, ByVal s As Single)
Public Declare Sub Vec4CatmullRom Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV0 As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4, ByRef pV3 As D3DVECTOR4, ByVal s As Single)
Public Declare Sub Vec4BaryCentric Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV1 As D3DVECTOR4, ByRef pV2 As D3DVECTOR4, ByRef pV3 As D3DVECTOR4, ByVal f As Single, ByVal G As Single)
Public Declare Sub Vec4Transform Lib "dx9_vb" (ByRef pOut As D3DVECTOR4, ByRef pV As D3DVECTOR4, ByRef pM As D3DMATRIX)
Public Declare Sub MatrixIdentity Lib "dx9_vb" (ByRef pOut As D3DMATRIX)
Public Declare Function MatrixIsIdentity Lib "dx9_vb" (ByRef pM As D3DMATRIX) As Boolean
Public Declare Function MatrixDeterminant Lib "dx9_vb" (ByRef pM As D3DMATRIX) As Single
Public Declare Sub MatrixTranspose Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pM As D3DMATRIX)
Public Declare Sub MatrixMultiply Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pM1 As D3DMATRIX, ByRef pM2 As D3DMATRIX)
Public Declare Sub MatrixMultiplyTranspose Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pM1 As D3DMATRIX, ByRef pM2 As D3DMATRIX)
Public Declare Sub MatrixInverse Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pDeterminant As Single, ByRef pM As D3DMATRIX)
Public Declare Sub MatrixScaling Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal sx As Single, ByVal sy As Single, ByVal sz As Single)
Public Declare Sub MatrixTranslation Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub MatrixRotationX Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal angle As Single)
Public Declare Sub MatrixRotationY Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal angle As Single)
Public Declare Sub MatrixRotationZ Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal angle As Single)
Public Declare Sub MatrixRotationAxis Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pV As D3DVECTOR, ByVal angle As Single)
Public Declare Sub MatrixRotationQuaternion Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pQ As D3DQUATERNION)
Public Declare Sub MatrixRotationYawPitchRoll Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal Yaw As Single, ByVal Pitch As Single, ByVal Roll As Single)
Public Declare Sub MatrixTransformation Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pScalingCenter As D3DVECTOR, ByRef pScalingRotation As D3DQUATERNION, ByRef pScaling As D3DVECTOR, ByRef pRotationCenter As D3DVECTOR, ByRef pRotation As D3DQUATERNION, ByRef pTranslation As D3DVECTOR)
Public Declare Sub MatrixAffineTransformation Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal Scaling As Single, ByRef pRotationCenter As D3DVECTOR, ByRef pRotation As D3DQUATERNION, ByRef pTranslation As D3DVECTOR)
Public Declare Sub MatrixLookAtRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pEye As D3DVECTOR, ByRef pAt As D3DVECTOR, ByRef pUp As D3DVECTOR)
Public Declare Sub MatrixLookAtLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pEye As D3DVECTOR, ByRef pAt As D3DVECTOR, ByRef pUp As D3DVECTOR)
Public Declare Sub MatrixPerspectiveRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal w As Single, ByVal h As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixPerspectiveLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal w As Single, ByVal h As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixPerspectiveFovRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal FovY As Single, ByVal Aspect As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixPerspectiveFovLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal FovY As Single, ByVal Aspect As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixPerspectiveOffCenterRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixPerspectiveOffCenterLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixOrthoRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal w As Single, ByVal h As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixOrthoLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal w As Single, ByVal h As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixOrthoOffCenterRH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixOrthoOffCenterLH Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub MatrixShadow Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pLight As D3DVECTOR4, ByRef pPlane As D3DPLANE)
Public Declare Sub MatrixReflect Lib "dx9_vb" (ByRef pOut As D3DMATRIX, ByRef pPlane As D3DPLANE)
Public Declare Function QuaternionLength Lib "dx9_vb" (ByRef pQ As D3DQUATERNION) As Single
Public Declare Function QuaternionLengthSq Lib "dx9_vb" (ByRef pQ As D3DQUATERNION) As Single
Public Declare Function QuaternionDot Lib "dx9_vb" (ByRef pQ1 As D3DQUATERNION, ByRef pQ2 As D3DQUATERNION) As Single
Public Declare Sub QuaternionIdentity Lib "dx9_vb" (ByRef pOut As D3DQUATERNION)
Public Declare Function QuaternionIsIdentity Lib "dx9_vb" (ByRef pQ As D3DQUATERNION) As Boolean
Public Declare Sub QuaternionConjugate Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ As D3DQUATERNION)
Public Declare Sub QuaternionToAxisAngle Lib "dx9_vb" (ByRef pQ As D3DQUATERNION, ByRef pAxis As D3DVECTOR, ByRef pAngle As Single)
Public Declare Sub QuaternionRotationMatrix Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pM As D3DMATRIX)
Public Declare Sub QuaternionRotationAxis Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pV As D3DVECTOR, ByVal angle As Single)
Public Declare Sub QuaternionRotationYawPitchRoll Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByVal Yaw As Single, ByVal Pitch As Single, ByVal Roll As Single)
Public Declare Sub QuaternionMultiply Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ1 As D3DQUATERNION, ByRef pQ2 As D3DQUATERNION)
Public Declare Sub QuaternionNormalize Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ As D3DQUATERNION)
Public Declare Sub QuaternionInverse Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ As D3DQUATERNION)
Public Declare Sub QuaternionLn Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ As D3DQUATERNION)
Public Declare Sub QuaternionExp Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ As D3DQUATERNION)
Public Declare Sub QuaternionSlerp Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ1 As D3DQUATERNION, ByRef pQ2 As D3DQUATERNION, ByVal t As Single)
Public Declare Sub QuaternionSquad Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ1 As D3DQUATERNION, ByRef pA As D3DQUATERNION, ByRef pB As D3DQUATERNION, ByRef pC As D3DQUATERNION, ByVal t As Single)
Public Declare Sub QuaternionSquadSetup Lib "dx9_vb" (ByRef pAOut As D3DQUATERNION, ByRef pBOut As D3DQUATERNION, ByRef pCOut As D3DQUATERNION, ByRef pQ0 As D3DQUATERNION, ByRef pQ1 As D3DQUATERNION, ByRef pQ2 As D3DQUATERNION, ByRef pQ3 As D3DQUATERNION)
Public Declare Sub QuaternionBaryCentric Lib "dx9_vb" (ByRef pOut As D3DQUATERNION, ByRef pQ1 As D3DQUATERNION, ByRef pQ2 As D3DQUATERNION, ByRef pQ3 As D3DQUATERNION, ByVal f As Single, ByVal G As Single)
Public Declare Function PlaneDot Lib "dx9_vb" (ByRef PP As D3DPLANE, ByRef pV As D3DVECTOR4) As Single
Public Declare Function PlaneDotCoord Lib "dx9_vb" (ByRef PP As D3DPLANE, ByRef pV As D3DVECTOR) As Single
Public Declare Function PlaneDotNormal Lib "dx9_vb" (ByRef PP As D3DPLANE, ByRef pV As D3DVECTOR) As Single
Public Declare Sub PlaneNormalize Lib "dx9_vb" (ByRef pOut As D3DPLANE, ByRef PP As D3DPLANE)
Public Declare Sub PlaneIntersectLine Lib "dx9_vb" (ByRef pOut As D3DVECTOR, ByRef PP As D3DPLANE, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR)
Public Declare Sub PlaneFromPointNormal Lib "dx9_vb" (ByRef pOut As D3DPLANE, ByRef pPoint As D3DVECTOR, ByRef pNormal As D3DVECTOR)
Public Declare Sub PlaneFromPoints Lib "dx9_vb" (ByRef pOut As D3DPLANE, ByRef pV1 As D3DVECTOR, ByRef pV2 As D3DVECTOR, ByRef pV3 As D3DVECTOR)
Public Declare Sub PlaneTransform Lib "dx9_vb" (ByRef pOut As D3DPLANE, ByRef PP As D3DPLANE, ByRef pM As D3DMATRIX)
Public Declare Sub ColorNegative Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC As D3DCOLORVALUE)
Public Declare Sub ColorAdd Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC1 As D3DCOLORVALUE, ByRef pC2 As D3DCOLORVALUE)
Public Declare Sub ColorSubtract Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC1 As D3DCOLORVALUE, ByRef pC2 As D3DCOLORVALUE)
Public Declare Sub ColorScale Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC As D3DCOLORVALUE, ByVal s As Single)
Public Declare Sub ColorModulate Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC1 As D3DCOLORVALUE, ByRef pC2 As D3DCOLORVALUE)
Public Declare Sub ColorLerp Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC1 As D3DCOLORVALUE, ByRef pC2 As D3DCOLORVALUE, ByVal s As Single)
Public Declare Sub ColorAdjustSaturation Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC As D3DCOLORVALUE, ByVal s As Single)
Public Declare Sub ColorAdjustContrast Lib "dx9_vb" (ByRef pOut As D3DCOLORVALUE, ByRef pC As D3DCOLORVALUE, ByVal c As Single)
Public Declare Function D3DCOLOR Lib "dx9_vb" Alias "ARGB_" (ByVal a As Byte, ByVal r As Byte, ByVal G As Byte, ByVal b As Byte) As Long
Public Declare Function D3DColorValueFromColor Lib "dx9_vb" Alias "D3DColorValue_" (ByVal c As Long) As D3DCOLORVALUE
Private Declare Sub d3d_Create Lib "dx9_vb" (ByRef pD3d As Long)
Private Ri              As Double
Public WantJump         As Boolean
Public TargetSpeed      As D3DVECTOR
Public oldis            As XINPUT_STATE
Public OldTime          As Double
Public NowTime          As Double
Public PlPos            As D3DVECTOR
Public PlSpeed          As D3DVECTOR
Public PlForce          As D3DVECTOR
Public PlDir            As D3DVECTOR
Public PlUp             As D3DVECTOR
Public PlRight          As D3DVECTOR
Public PlAngle          As Single
Public PlDiff           As Single
Public PlIsFly          As Boolean
Public PlIsCrouch       As Boolean
Public PlHeight         As Single

Public PlFlFrc          As D3DVECTOR
Public MaxSpeed         As Single
Public AddPos           As D3DVECTOR

Public Door()           As tDoor
Public DoorCnt          As Long
Public But()            As tButton
Public ButCnt           As Long
Public My               As New My
Private Type tDoor
    Mtrx As D3DMATRIX
    Sect As Long
    Tag As Long
    yMin As Single
    yMax As Single
    pMin As Single
    pMax As Single
    PosH As Single
    Spd As Single
    Pos As D3DVECTOR
End Type
Private Type tButton
    Mtrx As D3DMATRIX
    Sect As Long
    Target As Long
    Pos As D3DVECTOR
    Norm As D3DVECTOR
    Col As D3DVECTOR4
End Type
Public cCnt As Long
'Public ProjectileCnt As Long
Private Type tProjectiles
    Pos As D3DVECTOR
    Speed As D3DVECTOR
    Timer As Long
End Type
Public projectileselecting As Boolean
Public Type tCloudT
    Spd As D3DVECTOR
    Timer As Long
End Type
Public CloudT()    As tCloudT
'Public Const MaxProjectileCnt            As Long = 1
Public Projectiles As tProjectiles '(MaxProjectileCnt - 1) As tProjectiles
Public Thing()     As tThing
Public ThingPath() As String
Public ThingCnt    As Long
Public Type tThing
    InitPos As D3DVECTOR
    Pos As D3DVECTOR
    Mtrx As D3DMATRIX
    MeshInd As Long
End Type
Public Type Map3Vertex
    Pos As D3DVECTOR
    tu As Single
    tv As Single
    Tang As D3DVECTOR
    Bnrm As D3DVECTOR
    Norm As D3DVECTOR
End Type
Public Type tThing2D
    x As Integer
    z As Integer
    y As Integer
    tType As Byte
    Tag As Byte
End Type
Private Type VSConstants
    WVP As D3DMATRIX          ' 0
    World As D3DMATRIX        ' 4
    LitePos As D3DVECTOR      ' 8
    LitePosW As Single        '   = 0
    CamPos As D3DVECTOR       ' 9
    CamPosW As Single         '
    Diffuse As D3DCOLORVALUE  '10
    Specular As D3DCOLORVALUE '11
    mShad As D3DMATRIX        '12
End Type
Private Type PSConstants
    Ambient As D3DVECTOR4
    FogVal As D3DVECTOR4
    FogColor As D3DVECTOR4
End Type
Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type
Public Enum TextureFilter
    TextureFilter_BiLinear = 0
    TextureFilter_TriLinear = 1
    TextureFilter_Anisotropic = 2
End Enum
Public MapFileName   As String
Public LS            As clsLandScape
Public StartPos      As D3DVECTOR
Public SectorCnt     As Long
Public Sector()      As clsSector
Public ZOptimize     As Boolean
Public TexAlign()    As Long
Public TexScale()    As Single
Public RTTex         As Direct3DTexture9
Public mViewShad     As D3DMATRIX
Public mProjShad     As D3DMATRIX
Public SunLightPos   As D3DVECTOR
Public Dev           As Direct3DDevice9
Public mView         As D3DMATRIX
Public mProj         As D3DMATRIX
Public VSConst       As VSConstants
Public PSConst       As PSConstants
Public RenderEnabled As Boolean
Public BBWidth       As Long
Public BBHeight      As Long
Public directxTexFilters                  As TextureFilter
Public directxVSync As Boolean
Public directxQuant As Double
Public directxGravity As Single
Public directxTexFIndex As Long
Public directxAnisotropy As Long
Public directxFovY As Single
Public directxAspect As Single

Public keymapA As String
Public keymapB As String
Public keymapX As String
Public keymapY As String
Public keymapLeftBumper As String
Public keymapRightBumper As String
Public keymapDLeft As String
Public keymapDRight As String
Public keymapDUp As String
Public keymapDDown As String
Public keymapLeftStick As String
Public keymapRightStick As String
Public keymapMenu As String
Public keymapChange As String
Public keymapLThumbUp As Long
Public keymapLThumbDown As Long
Public keymapLThumbLeft As Long
Public keymapLThumbRight As Long
Public keymapRThumbUp As Long
Public keymapRThumbDown As Long
Public keymapRThumbLeft As Long
Public keymapRThumbRight As Long
Public keymapDisablegamepad As Long
Public pointerMaxPointerSpeed As Long
Public pointerMaxPointerAcceleration As Long
Public pointerMaxWheelSpeed As Long
Public pointerMaxWheelAcceleration As Long
Public pointerMaxPOVSpeed As Long
Public pointerMaxPOVAcceleration As Long
Public pointerMaxWalkSpeed As Long
Public pointerMaxWalkAcceleration As Long
Public pointerDisable2D As Long
Public pointerDisable3D As Long
Public displayHide As Long
Public displayDelay As Long
Public displaySpeed As Long
Public displayFade As Long
Public displayPosition As Long
Public displayTrans3D As Long
Public displayTransSettings As Integer
Public soundxMute As Long

Public MSGRender      As Boolean
Private Const GW As Long = (-20)
Private Const LWA_COLORKEY As Long = &H1        'to trans'
Private Const LWA_ALPHA As Long = &H2           'to semi trans'
Private Const WS_EX_LAYERED  As Long = &H80000
Private Declare Function apiSetLayeredWindowAttributes Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As VbAppWinStyle) As Long
Public folderinview As String
Public tex()        As Direct3DTexture9

Private Sub Main()
    On Error Resume Next
    '    Dim ds As New clsDriveSerial
    '    If ds.GetDesktopName = "Default" Then
    '      If App.PrevInstance = True Then End
    '    End If
    '    If ds.IsDiskDrivePresent = False Then End
    If Len(Command$()) > 0 Then
        If InStr(1, LCase(Split(Command$, " ")(0)), "runas") <> 0 Then
            'InitCommonControls
            frmMain.show
        End If
    Else
        Dim Params As Variant
        Params = Array("runas")
        ShellExecute 0, "runas", App.Path & "\" & App.EXEName & ".exe", Join(Params, " "), CurDir$(), vbNormalFocus
    End If
End Sub
Public Sub CreateRoundRectFromWindow(ByRef oWindow As Object)
    On Error Resume Next
    Dim lRight  As Long
    Dim lBottom As Long
    Dim hRgn    As Long
    With oWindow
        lRight = .Width / Screen.TwipsPerPixelX
        lBottom = .Height / Screen.TwipsPerPixelY
        hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 20, 20)
        SetWindowRgn .hWnd, hRgn, True
    End With
End Sub
Public Sub CreateRoundRectFromWindow2(ByRef oWindow As Object)
    On Error Resume Next
    Dim lRight  As Long
    Dim lBottom As Long
    Dim hRgn    As Long
    With oWindow
        lRight = .Width / Screen.TwipsPerPixelX
        lBottom = .Height / Screen.TwipsPerPixelY
        hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 131, 131)
        SetWindowRgn .hWnd, hRgn, True
    End With
End Sub
Public Sub CreateRoundRectFromcontrol(ByRef oWindow As Object)
    On Error Resume Next
    Dim lRight  As Long
    Dim lBottom As Long
    Dim hRgn    As Long
    With oWindow
        lRight = .Width / Screen.TwipsPerPixelX
        lBottom = .Height / Screen.TwipsPerPixelY
        hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 5, 5)
        SetWindowRgn .hWnd, hRgn, True
    End With
End Sub
Public Sub WindowTransparency(ByVal hWnd As Long, ByVal Level As Integer, ByVal Color As Long)
    On Error Resume Next
    Dim Msg As Long
    Msg = apiGetWindowLong(hWnd, GW)
    Msg = Msg Or WS_EX_LAYERED
    apiSetWindowLong hWnd, GW, Msg
    apiSetLayeredWindowAttributes hWnd, Color, Level, LWA_ALPHA
End Sub
Public Function Rand() As Single
    Ri = 1.314 * Ri + 1.737
    If Ri > 983732.3456 Then Ri = Ri * 0.3141
    Rand = Ri - Int(Ri)
End Function
Private Sub RandInit(r As Single)
    Ri = r
End Sub
Public Function PlIntersect(v As D3DVECTOR, Optional ByVal Rad As Single = 0) As Boolean
    If v.y < PlPos.y + HeadHeight + Rad And v.y > PlPos.y - PlHeight - Rad Then
        If (PlPos.x - v.x) * (PlPos.x - v.x) + (PlPos.z - v.z) * (PlPos.z - v.z) < 2 + Rad * Rad Then PlIntersect = True
    End If
End Function
Public Function GetImageInfoFromFile(fName As String) As D3DXIMAGE_INFO
    Dim imgInfo As D3DXIMAGE_INFO
    D3DX_GetImageInfoFromFile StrPtr(fName), imgInfo
    GetImageInfoFromFile = imgInfo
End Function
Public Function CreateTextureFromFile(d3dDev As Direct3DDevice9, fName As String) As Direct3DTexture9
    Dim pTex As Long
    D3DX_CreateTextureFromFile d3dDev.Ptr, StrPtr(fName), pTex
    If pTex <> 0 Then
        Set CreateTextureFromFile = New Direct3DTexture9
        CreateTextureFromFile.Ptr = pTex
    End If
End Function
Public Function CreateTextureFromFileEx(d3dDev As Direct3DDevice9, fName As String, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As D3DUSAGE, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As D3DX_FILTER, ByVal MipFilter As D3DX_FILTER, ByVal ColorKey As Long) As Direct3DTexture9
    Dim pTex As Long
    D3DX_CreateTextureFromFileEx d3dDev.Ptr, StrPtr(fName), Width, Height, MipLevels, Usage, Format, Pool, Filter, MipFilter, ColorKey, 0, 0, pTex
    If pTex <> 0 Then
        Set CreateTextureFromFileEx = New Direct3DTexture9
        CreateTextureFromFileEx.Ptr = pTex
    End If
End Function
Public Function CreateDirectSound() As DirectSound8
    Dim pDS As Long
    ds_Create pDS
    If pDS <> 0 Then
        Set CreateDirectSound = New DirectSound8
        CreateDirectSound.Ptr = pDS
    End If
End Function
Public Function D3DFVF_TEXCOORDSIZE(ByVal CoordIndex As Long, ByVal Size As Long) As CONST_D3DFVF
    Dim d As Long
    d = (Size + 2) And 3
    D3DFVF_TEXCOORDSIZE = d * 2 ^ (CoordIndex * 2 + 16)
End Function
Public Function VertexElement(Stream As Integer, Offset As Integer, dType As D3DDECLTYPE, Method As D3DDECLMETHOD, Usage As D3DDECLUSAGE, UsageIndex As Long) As D3DVERTEXELEMENT9
    VertexElement.Stream = Stream
    VertexElement.Offset = Offset
    VertexElement.dType = dType
    VertexElement.Method = Method
    VertexElement.Usage = Usage
    VertexElement.UsageIndex = UsageIndex
End Function
Public Function VertexElementEnd() As D3DVERTEXELEMENT9
    VertexElementEnd = VertexElement(&HFF, 0, D3DDECLTYPE_UNUSED, 0, 0, 0)
End Function
Public Function D3DTS_WORLDMATRIX(ByVal index As Long) As D3DTRANSFORMSTATETYPE
    D3DTS_WORLDMATRIX = index + 256
End Function
Public Function CreateDirect3D() As Direct3D9
    Dim pD3d As Long
    d3d_Create pD3d
    If pD3d <> 0 Then
        Set CreateDirect3D = New Direct3D9
        CreateDirect3D.Ptr = pD3d
    End If
End Function
Public Function Vec2(ByVal x As Single, ByVal y As Single) As D3DVECTOR2
    Vec2.x = x
    Vec2.y = y
End Function
Public Function Vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    Vec3.x = x
    Vec3.y = y
    Vec3.z = z
End Function
Public Function Vec4(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single) As D3DVECTOR4
    Vec4.x = x
    Vec4.y = y
    Vec4.z = z
    Vec4.w = w
End Function
Public Function D3DCOLORVALUE(ByVal r As Single, ByVal G As Single, ByVal b As Single, ByVal a As Single) As D3DCOLORVALUE
    D3DCOLORVALUE.r = r
    D3DCOLORVALUE.G = G
    D3DCOLORVALUE.b = b
    D3DCOLORVALUE.a = a
End Function
Public Function D3DColorFromColorValue(c As D3DCOLORVALUE) As Long
    D3DColorFromColorValue = D3DCOLOR(c.a, c.r, c.G, c.b)
End Function
'Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
'Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private sysSurf     As Direct3DSurface9
'Private d3dev As Direct3DDevice9
'Private d3d9 As Direct3D9
'Private biWnd       As BITMAPINFO
'Private vtxBuf      As Direct3DVertexBuffer9
'Private surf        As Direct3DSurface9                    ' // Surface for render to texture
'Private texture     As Direct3DTexture9                    ' // Texture of window
'Private backTex     As Direct3DTexture9                    ' // Texture of render target
'Private isModify    As Boolean                  ' // If this flag set then window is rotated
'Private triggerXMin As Long                     ' // Minimum positions when the rotate is enabled
'Private triggerXMax As Long
'Private triggerYMin As Long
'Private triggerYMax As Long
'Private bmpShadow   As Long                     ' // Bitmap in memory, which represent the window
'Private lpBmpData   As Long                     ' // Pointer to the bmpShadow bits
'Private isInit      As Boolean
'
'
'Private Sub InitializeRotateWindow(ByVal hdc As Long, ByVal w As Long, ByVal h As Long)
'    Dim sWidth  As Long
'    Dim sHeight As Long
'    Dim pP      As D3DPRESENT_PARAMETERS
'    biWnd.bmiHeader.biSize = Len(biWnd.bmiHeader)
'    biWnd.bmiHeader.biBitCount = 32
'    biWnd.bmiHeader.biHeight = -h / Screen.TwipsPerPixelY
'    biWnd.bmiHeader.biWidth = w / Screen.TwipsPerPixelY
'    biWnd.bmiHeader.biPlanes = 1
'    bmpShadow = CreateDIBSection(hdc, biWnd, 0, lpBmpData, 0, 0)
'    sWidth = Screen.Width / Screen.TwipsPerPixelX - biWnd.bmiHeader.biWidth
'    sHeight = Screen.Height / Screen.TwipsPerPixelY + biWnd.bmiHeader.biHeight
'    triggerXMin = sWidth * (1 / 5)
'    triggerXMax = sWidth - triggerXMin
'    triggerYMin = sHeight * (1 / 5)
'    triggerYMax = sHeight - triggerYMin
'    Set d3d9 = CreateDirect3D()
'    pP.BackBufferCount = 1
'    pP.Windowed = 1
'    pP.BackBufferFormat = D3DFMT_A8R8G8B8
'    pP.SwapEffect = D3DSWAPEFFECT_DISCARD
'    pP.EnableAutoDepthStencil = 1
'    pP.AutoDepthStencilFormat = D3DFMT_D16
'    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, pP)
'    Dim vtx()   As Vertex
'    Dim lpDat   As Long
'    ReDim vtx(5)
'    nPlan Vec3(-biWnd.bmiHeader.biWidth, biWnd.bmiHeader.biHeight, 0), Vec3(biWnd.bmiHeader.biWidth, biWnd.bmiHeader.biHeight, 0), Vec3(biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 0), Vec3(-biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 0), 0, vtx(), 0, 0, 1, 1
'
'    d3dev.CreateVertexBuffer Len(vtx(0)) * (UBound(vtx) + 1), D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_TEX1, D3DPOOL_DEFAULT, vtxBuf
'
'    vtxBuf.Lock 0, 0, lpDat, 0
'    memcpy ByVal lpDat, vtx(0), Len(vtx(0)) * (UBound(vtx) + 1)
'    vtxBuf.Unlock
'    d3dev.SetFVF D3DFVF_XYZ Or D3DFVF_TEX1
'    d3dev.SetStreamSource 0, vtxBuf, 0, 5 * 4
'    Dim Mtx As D3DMATRIX
'    Dim fov As Single
'    Dim l   As Single
'    fov = Pi / 3
'    ' // Calculate distance to billboard in order to fit window to the render area
'    l = -biWnd.bmiHeader.biHeight * Tan(fov)
'    D3DXMatrixLookAtLH Mtx, Vec3(0, 0, -l), Vec3(0, 0, 0), Vec3(0, 1, 0)
'    d3dev.SetTransform D3DTS_VIEW, Mtx
'    D3DXMatrixPerspectiveFovLH Mtx, fov, Width / Height, 1, 10000
'    d3dev.SetTransform D3DTS_PROJECTION, Mtx
'    ' // Create the window-texture
'    d3dev.CreateTexture biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 1, D3DUSAGE_DYNAMIC, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, texture
'    ' // Create the render-target texture
'    d3dev.CreateTexture biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, backTex
'    ' // Create the lockable surface
'    d3dev.CreateOffscreenPlainSurface biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, D3DFMT_A8R8G8B8, D3DPOOL_SYSTEMMEM, sysSurf
'    Set surf = backTex.GetSurfaceLevel(0)
'    d3dev.SetRenderTarget 0, surf
'    d3dev.SetTexture 0, texture
'    d3dev.SetRenderState D3DRS_LIGHTING, 0
'    d3dev.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
'    d3dev.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
'    d3dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
'    d3dev.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_CONSTANT
'    d3dev.SetTextureStageState 0, D3DTSS_CONSTANT, &HFF000000
'    isInit = True
'End Sub
'Private Sub RotateWindow(ByVal dx As Single, ByVal dy As Single)
'    Dim Mtx As D3DMATRIX
'    Dim off As Single
'    '  Calc maximum offset
'    If Abs(dx) > Abs(dy) Then off = Abs(dx) Else off = Abs(dy)
'    '  Move aside the window
'   ' D3DXMatrixTranslation Mtx, 0, 0, off * 300
'    d3dev.SetTransform D3DTS_WORLD, Mtx
'    '  Rotation the window
'    D3DXMatrixRotationY Mtx, dx
'    d3dev.MultiplyTransform D3DTS_WORLD, Mtx
'    D3DXMatrixRotationX Mtx, dy
'    d3dev.MultiplyTransform D3DTS_WORLD, Mtx
'    d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
'    d3dev.BeginScene
'    d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
'    d3dev.EndScene
'    d3dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
'    Dim pt      As Size
'    Dim sz      As Size
'    Dim pos     As Size
'    Dim rect    As D3DLOCKED_RECT
'    pt.cx = Me.Left / Screen.TwipsPerPixelX
'    pt.cy = Me.Top / Screen.TwipsPerPixelY
'    sz.cx = biWnd.bmiHeader.biWidth
'    sz.cy = -biWnd.bmiHeader.biHeight
'    d3dev.GetRenderTargetData surf, sysSurf ' // Copy bitmap to the system memory surface
'    sysSurf.LockRect rect, ByVal 0&, D3DLOCK_DISCARD
'    ' Copy to form
'    SetDIBitsToDevice Me.hdc, 0, 0, sz.cx, sz.cy, 0, 0, 0, sz.cy, ByVal rect.pBits, biWnd, 0
'    UpdateLayeredWindow Me.hWnd, Me.hdc, pt, sz, Me.hdc, pos, 0, AB_32Bpp255, ULW_ALPHA
'    sysSurf.UnlockRect
'End Sub
'Public Sub D3DXMatrixRotationY(pOut As D3DMATRIX, angle As Single) ' // Builds a matrix that rotates around the y-axis.
'    Dim s   As Single
'    Dim c   As Single
'    D3DXMatrixIdentity pOut
'    s = Sin(angle)
'    c = Cos(angle)
'    pOut.m11 = c
'    pOut.m33 = c
'    pOut.m13 = -s
'    pOut.m31 = s
'End Sub
'
'
