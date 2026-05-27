Attribute VB_Name = "modScreenStatusFX"
' Argentum 20 Cliente - overlays de pantalla para estados visuales.
' Neurotoxina es el primer uso: vignette organico, sutil y no bloqueante.

Option Explicit

Private Const NEURO_EFFECT_TYPE_ID As Integer = 45
Private Const NEURO_VIGNETTE_FILE As String = "neurotoxina_vignette.png"
Private Const NEURO_VIGNETTE_SIZE As Long = 1024
Private Const NEURO_ALPHA_MIN As Long = 150
Private Const NEURO_ALPHA_MAX As Long = 210
Private Const NEURO_PULSE_PERIOD_MS As Long = 1800

Private NeuroVignetteTexture As Direct3DTexture8
Private NeuroVignetteLoadFailed As Boolean

Public Sub RenderScreenStatusFX()
    On Error GoTo RenderScreenStatusFX_Err

    Dim sw As Long
    Dim sh As Long

    sw = frmMain.renderer.ScaleWidth
    sh = frmMain.renderer.ScaleHeight

    If sw <= 0 Or sh <= 0 Then Exit Sub

    If IsNeuroActive() Then
        Call RenderNeurotoxinaOverlay(sw, sh)
    End If

    Exit Sub
RenderScreenStatusFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "modScreenStatusFX.RenderScreenStatusFX", Erl)
End Sub

Public Sub ReleaseScreenStatusFX()
    Set NeuroVignetteTexture = Nothing
    NeuroVignetteLoadFailed = False
End Sub

Private Function IsNeuroActive() As Boolean
    On Error GoTo IsNeuroActive_Err

    IsNeuroActive = False

    If DeBuffList.EffectCount <= 0 Then Exit Function

    Dim i As Integer
    For i = 0 To DeBuffList.EffectCount - 1
        If DeBuffList.EffectList(i).TypeId = NEURO_EFFECT_TYPE_ID Then
            IsNeuroActive = True
            Exit Function
        End If
    Next i

    Exit Function
IsNeuroActive_Err:
    Call RegistrarError(Err.Number, Err.Description, "modScreenStatusFX.IsNeuroActive", Erl)
End Function

Private Sub RenderNeurotoxinaOverlay(ByVal sw As Long, ByVal sh As Long)
    On Error GoTo RenderNeurotoxinaOverlay_Err

    If Not EnsureOverlayTexture(NEURO_VIGNETTE_FILE, NeuroVignetteTexture, NeuroVignetteLoadFailed) Then Exit Sub

    Dim colors(3) As RGBA
    Call RGBAList(colors, 255, 255, 255, PulseAlpha(NEURO_ALPHA_MIN, NEURO_ALPHA_MAX, NEURO_PULSE_PERIOD_MS))

    Call SpriteBatch.SetTexture(NeuroVignetteTexture)
    Call SpriteBatch.SetAlpha(False)
    Call SpriteBatch.Draw(0, 0, sw, sh, colors, 0, 0, 1, 1)

    Exit Sub
RenderNeurotoxinaOverlay_Err:
    Call RegistrarError(Err.Number, Err.Description, "modScreenStatusFX.RenderNeurotoxinaOverlay", Erl)
End Sub

Private Function EnsureOverlayTexture(ByVal fileName As String, ByRef textureRef As Direct3DTexture8, ByRef loadFailedRef As Boolean) As Boolean
    On Error GoTo EnsureOverlayTexture_Err

    If Not textureRef Is Nothing Then
        EnsureOverlayTexture = True
        Exit Function
    End If

    If loadFailedRef Then Exit Function

    #If Compresion = 1 Then
        Dim bytArr() As Byte
        If Not Extract_File_To_Memory(interface, App.path & "\..\Recursos\OUTPUT", fileName, bytArr(), ResourcesPassword) Then
            loadFailedRef = True
            frmDebug.add_text_tracebox "No se pudo cargar overlay de pantalla: " & fileName & " desde OUTPUT/Interface."
            Exit Function
        End If

        Set textureRef = DirectD3D8.CreateTextureFromFileInMemoryEx(DirectDevice, bytArr(0), UBound(bytArr) + 1, NEURO_VIGNETTE_SIZE, NEURO_VIGNETTE_SIZE, 1, 0, _
                D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
    #Else
        Dim path As String
        path = App.path & "\..\Recursos\Interface\" & fileName
        If Not FileExist(path, vbArchive) Then
            loadFailedRef = True
            frmDebug.add_text_tracebox "No se pudo cargar overlay de pantalla: " & path
            Exit Function
        End If

        Set textureRef = DirectD3D8.CreateTextureFromFileEx(DirectDevice, path, NEURO_VIGNETTE_SIZE, NEURO_VIGNETTE_SIZE, 1, 0, D3DFMT_A8R8G8B8, _
                D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
    #End If

    EnsureOverlayTexture = Not textureRef Is Nothing
    Exit Function
EnsureOverlayTexture_Err:
    loadFailedRef = True
    Call RegistrarError(Err.Number, Err.Description, "modScreenStatusFX.EnsureOverlayTexture", Erl)
End Function

Private Function PulseAlpha(ByVal minAlpha As Long, ByVal maxAlpha As Long, ByVal periodMs As Long) As Byte
    Dim phase As Single
    Dim pulse As Single

    If periodMs <= 0 Then
        PulseAlpha = CByte(maxAlpha)
        Exit Function
    End If

    phase = (CSng(GetTickCount() Mod periodMs) / CSng(periodMs)) * 6.2831853
    pulse = (Sin(phase) + 1) / 2
    PulseAlpha = CByte(minAlpha + CLng(pulse * (maxAlpha - minAlpha)))
End Function
