VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cGetTimer                                                         (16/Jun/98)
' Clase para calcular el tiempo entre eventos, etc.
'
' Ampliada para usar distintas funciones del API                    (29/Ago/98)
' Nuevos m�todos/propiedades usando APIs especiales para NT/2000    (27/Oct/00)
' Comprobado en Windows XP Professional                             (09/Jul/03)
'
' Para usarla:
'   StartTimer          Iniciar el temporizador
'   StopTimer           Finalizar la cuenta
'   ElapsedTimer        Mostrar el tiempo transcurrido, formateado
'                       --- Nuevas propiedades ---
'   TimerType           El temporizador a usar
'   TotalElapsedTime    El n�mero de milisegundos transcurridos
'
'   TimerPerformanceFrecuency
'   TimerPerformanceCounter
'
' �Guillermo 'guille' Som, 1998-2003
'
' En esta clase se usa el tiempo del sistema, es decir el n�mero de
' milisegundos transcurridos desde que se inici� Windows.
' Este temporizador vuelve a cero una vez transcurridos 49.7 d�as,
' (2^32 milisegundos),
' por tanto el c�lculo del tiempo transcurrido puede no ser exacto,
' dependiendo de cu�ndo se inici� y termin� la cuenta.
'------------------------------------------------------------------------------
Option Explicit

'
'------------------------------------------------------------------------------
' Se puede usar timeGetTime si se quiere mayor precisi�n,
' ya que esta funci�n tiene una precisi�n de:
' 1 milisegundo en Windows 95, y de
' 5 milisegundos o m�s, (es configurable), en Windows NT
'
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'
' S�lo para Windows NT/2000/XP
' Para averiguar si se dispone de temporizador de alta precisi�n
' y configurar la precisi�n de timeGetTime()
'
' Nota: El valor asignado a los par�metros Currency hay que multiplicarlo por 10000
Private Declare Function QueryPerformanceCounter Lib "kernel32" _
    (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
    (ByRef lpFrequency As Currency) As Long
'
' Texto basado en la descripci�n de:
'   High-Resolution Timer (Platform SDK / Windows User Interface)
' La funci�n QueryPerformanceCounter devuelve el valor actual del temporizador
' de alta precisi�n (si es que el sistema tiene uno).
' Se puede usar para calcular intervalos de tiempo usando la precisi�n de
' un temporizador de alta precisi�n (o resoluci�n).
' Por ejemplo, suponiendo que QueryPerformanceFrequency indicara que la frecuencia
' del temporizador de alta precisi�n es de 50000 por segundo. Si la aplicaci�n
' llamara a QueryPerformanceCounter inmediatamente antes y justo despu�s de la
' secci�n de c�digo a temporizar, y que estos valores fuesen 1500 y 3500,
' esto indicar�a que el tiempo transcurrido es de .04 segundos (2000 ciclos)
'
'
'Private Declare Function timeBeginPeriod Lib "winmm.dll" _
'    (ByVal uPeriod As Long) As Long
'Private Declare Function timeEndPeriod Lib "winmm.dll" _
'    (ByVal uPeriod As Long) As Long
'
'------------------------------------------------------------------------------
' La precisi�n de GetTickCount es (aproximadamente)
' de 10 milisegundos en Windows NT 3.5 o superior
' de 16 ms. en Windows NT 3.1
' de 55 ms. en Windows 95 o superior
'
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'------------------------------------------------------------------------------
' Valores del tipo de timer a usar
Public Enum eTipoTimer
    eGetTickCount '= 0
    eTimeGetTime '= 1
End Enum

' Variables locales para los c�lculos
'
Private mc_StartTimer As Currency
Private mc_EndTimer As Currency
Private mc_Frecuencia As Currency
'
Private m_TimerType As eTipoTimer
Private m_StartTimer As Long
Private m_EndTimer As Long
Private m_LastTime As Long      ' �ltimo tiempo calculado           (11/Mar/99)

Public Function ElapsedTime(Optional Intervalo As Long = 1000) As String
    ' Mostrar el tiempo transcurrido
    '
    ' Intervalo se usar� para la forma de mostrar el tiempo transcurrido
    '
    ' Detener el timer
    If m_EndTimer = 0 Then
        StopTimer
    End If
    '
    ' No admitir n�meros negativos ni cero
    If Intervalo < 1 Then Intervalo = 1000
    '
    ' Devolver una cadena formateada con el tiempo
    ElapsedTime = Format$(CDbl(m_EndTimer) / Intervalo, "0.0####")
    ' si se quiere usar el TimerPerformanceCounter
    'ElapsedTime = Format$((mc_EndTimer / mc_Frecuencia) * 1000 / Intervalo, "0.0####")
End Function

Public Sub StopTimer()
    ' Detener el timer
    '
    ' El tiempo de diferencia puede no ser "exacto",
    ' ya que se reinicia cada 49.7 dias... (2^32 milisegundos)
    ' El timer se pone a cero cuando se inicia el sistema.
    If m_TimerType = eTimeGetTime Then
        'mc_EndTimer = TimerPerformanceCounter - mc_StartTimer
        m_EndTimer = timeGetTime() - m_StartTimer
    Else
        m_EndTimer = GetTickCount() - m_StartTimer
    End If
    '
    m_LastTime = m_EndTimer
End Sub

Public Sub StartTimer()
    ' Iniciar el timer
    m_EndTimer = 0
    If m_TimerType = eTimeGetTime Then
        'mc_Frecuencia = TimerPerformanceFrecuency
        'mc_StartTimer = TimerPerformanceCounter
        m_StartTimer = timeGetTime()
    Else
        m_StartTimer = GetTickCount()
    End If
End Sub

Public Property Get TimerType() As eTipoTimer
    ' Devolver el tipo de timer que se usa
    TimerType = m_TimerType
End Property

Public Property Let TimerType(ByVal vNewValue As eTipoTimer)
    ' asignar el tipo de timer a usar
    m_TimerType = vNewValue
End Property

Private Sub Class_Initialize()
    ' Valor por defecto
    'm_TimerType = eGetTickCount
    m_TimerType = eTimeGetTime      ' Este es m�s preciso               (11/Mar/99)
End Sub

Public Property Get TotalElapsedTime() As Long
    ' Mostrar el tiempo transcurrido en el �ltimo c�lculo,
    ' incluso mientras se est� calculando
    TotalElapsedTime = m_LastTime ' m_EndTimer
End Property

Public Property Get TimerPerformanceFrecuency() As Long
    ' Devuelve la frecuencia del temporizador de alta precisi�n     (27/Oct/00)
    '
    Dim curFrecuency As Currency
    Dim ret As Long
    '
    ret = QueryPerformanceFrequency(curFrecuency)
    If ret = 0 Then
        ' No hay un temporizador de alta precisi�n
        TimerPerformanceFrecuency = 0
    Else
        ' Devolver el valor de la frecuencia por cada segundo
        ' (multiplicar por 10000 ya que el valor devuelto viene con decimales)
        TimerPerformanceFrecuency = curFrecuency * 10000
    End If
End Property

Public Property Get TimerPerformanceCounter() As Currency
    ' Contador del temporizador de alta precisi�n actual            (27/Oct/00)
    ' Esta funci�n se usar� junto con TimerPerformanceFrecuency,
    ' para averiguar el lapso transcurrido entre dos periodos de tiempo
    '
    ' Los milisegundos transcurridos ser�:
    '   la diferencia entre dos llamadas a esta propiedad
    '   dividido por el valor indicado por TimerPerformanceFrecuency
    '
    ' Nota: Este resultado ser� casi igual que una llamada a timeGetTime,
    '       aunque algo m�s precisa.
    '
    Dim curFrecuency As Currency
    Dim ret As Long
    '
    ret = QueryPerformanceCounter(curFrecuency)
    If ret = 0 Then
        ' No hay un temporizador de alta precisi�n
        TimerPerformanceCounter = 0
    Else
        ' Devolver el valor del contador
        ' (multiplicar por 10000 ya que el valor devuelto viene con decimales)
        TimerPerformanceCounter = curFrecuency * 10000
    End If
End Property
