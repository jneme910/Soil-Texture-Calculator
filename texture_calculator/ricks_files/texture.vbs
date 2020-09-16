
MSGBOX (OUTPUT())


FUNCTION OUTPUT() 
 DIM Clay
 DIM Silt
 DIM Sand
 DIM vfs
 DIM fs
 DIM ms
 DIM cs
 DIM vcs
 Dim Sil2Clay      ' silt plus 2 times the clay
 Dim SiltClay      ' silt plus 1.5 times the clay
 Dim SumVCSCS      ' the sum of vcs and cs
 Dim SVCSCSMS      ' the sum of vcs, cs, and ms
 Dim SumFSVFS      ' the sum of fs and vfs
 Dim RatFSVFS      ' the sum of fs and vfs divided by 2
 Dim Texture

' Clay = [clay_tot_psa]
' Silt = [silt_tot_psa]
' Sand = [sand_tot_psa]
' vfs = [sand_vf_psa]
' fs = [sand_f_psa]
' ms = [sand_m_psa]
' cs = [sand_c_psa]
' vcs = [sand_vc_psa]
'VC sand 1.0, 
'C sand 6.8, 
'medium sands 25.9, 
'F sands 52.3, 
'VF sands 6.9, 

'total sands 92.9, and 

'clay 4.6.  

 Clay = 4.6
 Silt = 2.5
 Sand = 92.9
 vfs = 6.9
 fs = 52.3
 ms = 25.9
 cs = 6.8
 vcs = 1.0

 IF Not IsNumeric(Clay) THEN Clay = 0.0
 IF Not IsNumeric(Silt) THEN Silt = 0.0
 IF Not IsNumeric(Sand) THEN Sand = 0.0
 IF Not IsNumeric(vfs) THEN vfs = 0.0
 IF Not IsNumeric(fs) THEN fs = 0.0
 IF Not IsNumeric(ms) THEN ms = 0.0
 IF Not IsNumeric(cs) THEN cs = 0.0
 IF Not IsNumeric(vcs) THEN vcs = 0.0

  Clay=AsymArithRnd(clay,10)
  vcs=AsymArithRnd(vcs,10)
  cs=AsymArithRnd(cs,10)
  fs=AsymArithRnd(fs,10)
  vfs=AsymArithRnd(vfs,10)
  ms=AsymArithRnd(ms,10)
  Silt=AsymArithRnd(Silt,10)
  Sand=AsymArithRnd(Sand,10)
 SiltClay = AsymArithRnd((Silt + (1.5 * Clay)), 10)
 SumVCSCS = AsymArithRnd((vcs + cs), 10)
 SVCSCSMS = AsymArithRnd((vcs + cs + ms), 10)
 Sil2Clay = AsymArithRnd((Silt + (2 * Clay)), 10)
 SumFSVFS = AsymArithRnd((fs + vfs), 10)
 RatFSVFS = AsymArithRnd(((fs + vfs) / 2), 10)

Texture = ""


 SELECT CASE True

  CASE AsymArithRnd((Sand + Silt + Clay),10) <> 100.0
        'Texture = Texture & ":" & "Error. Total ? 100. " & sand & "|" & silt & "|" & clay
	 Texture = Texture & ":" & "Error. Total ? 100."
'Sands
  CASE SiltClay < 15.0

Texture = Texture & "CASE SiltClay < 15.0" & vbcrlf

        IF SumVCSCS >= 25.0 AND ms < 50.0 AND fs < 50.0 AND vfs < 50.0 THEN
Texture = Texture & "IF SumVCSCS >= 25.0 AND ms < 50.0 AND fs < 50.0 AND vfs < 50.0" & vbcrlf
           Texture = Texture & ":" & "CoS"
        ELSEIF (SVCSCSMS >= 25.0 AND SumVCSCS <25.0 AND fs < 50.0 AND vfs < 50.0) OR  (SumVCSCS >= 25.0 And ms >= 50.0) THEN
Texture = Texture & "IF (SVCSCSMS >= 25.0 AND SumVCSCS <25.0 AND fs < 50.0 AND vfs < 50.0) OR  (SumVCSCS >= 25.0 And ms >= 50.0)" & vbcrlf
           Texture = Texture & ":" & "S"
	ElseIf (fs >= 50.0 And fs > vfs) Or (SVCSCSMS < 25.0 And vfs < 50.0) Then
Texture = Texture & "IF (fs >= 50.0 And fs > vfs) Or (SVCSCSMS < 25.0 And vfs < 50.0)" & vbcrlf
           Texture = Texture & ":" & "FS"
        ELSEIF vfs >= 50.0 THEN
Texture = Texture & "IF vfs >= 50.0" & vbcrlf
           Texture = Texture & ":" & "VFS"
        ELSE
Texture = Texture & "ELSE" & vbcrlf
           Texture = Texture & ":" & ""
        END IF
'Loamy Sands
  CASE  (SiltClay >= 15.0) AND ( Sil2Clay < 30.0)

Texture = Texture & "CASE (SiltClay >= 15.0) AND ( Sil2Clay < 30.0)" & vbcrlf

        IF SumVCSCS >= 25.0 AND ms < 50.0 AND fs < 50.0 AND vfs < 50.0 THEN
           Texture = Texture & ":" & "LCoS"
        ELSEIF (SVCSCSMS >= 25.0 And SumVCSCS < 25.0 AND fs < 50.0 AND vfs < 50.0) Or (SumVCSCS >= 25.0 And ms >= 50.0) THEN
           Texture = Texture & ":" & "LS"
        ELSEIF fs >= 50.0 OR (SVCSCSMS < 25.0 AND vfs < 50.0) THEN
           Texture = Texture & ":" & "LFS"
        ELSEIF vfs >= 50.0 THEN
           Texture = Texture & ":" & "LVFS"
        ELSE
           Texture = Texture & ":" & ""
        END IF
'Sandy Loams
  CASE (Clay >= 7.0 And Clay < 20.0 AND Sand > 52.0 AND Sil2Clay >= 30.0) OR (Clay < 7.0 AND Silt < 50.0 AND Sil2Clay >= 30.0)

Texture = Texture & "CASE (Clay >= 7.0 And Clay < 20.0 AND Sand > 52.0 AND Sil2Clay >= 30.0) OR (Clay < 7.0 AND Silt < 50.0 AND Sil2Clay >= 30.0)" & vbcrlf

    IF (SumVCSCS >= 25.0 AND ms < 50.0 AND fs < 50.0 AND vfs < 50.0) OR (SVCSCSMS >= 30.0 AND vfs >= 30.0 AND vfs < 50.0) THEN
           Texture = Texture & ":" & "CoSL"
    ELSEIF SVCSCSMS >= 30.0 And SumVCSCS < 25.0 AND vfs < 30.0 AND fs < 30.0 THEN
           Texture = Texture & ":" & "SL"
    ELSEIF (SVCSCSMS <= 15.0 AND vfs < 30.0 AND fs < 30.0 AND SumFSVFS < 40.0)  Or (SumVCSCS >= 25.0 And ms >= 50.0)  THEN
           Texture = Texture & ":" & "SL"
    ELSEIF fs >= 30.0 AND vfs < 30.0 And SumVCSCS < 25.0 THEN
           Texture = Texture & ":" & "FSL"
    ELSEIF SVCSCSMS >= 15.0 AND SVCSCSMS < 30.0 And SumVCSCS < 25.0  THEN
           Texture = Texture & ":" & "FSL"
    ELSEIF (SumFSVFS >= 40.0 And fs >= vfs And SVCSCSMS <= 15.0) Or (SumVCSCS >= 25.0 And fs >= 50.0) THEN
           Texture = Texture & ":" & "FSL"
    ELSEIF (vfs >= 30.0 And SVCSCSMS < 15.0 And vfs > fs) Or (SumFSVFS >= 40.0 And vfs > fs And SVCSCSMS < 15.0) THEN
           Texture = Texture & ":" & "VFSL"
    ELSEIF (vfs >= 50.0 And SumVCSCS >= 25.0) Or (vfs >= 50.0 And SVCSCSMS >= 30.0) THEN
           Texture = Texture & ":" & "VFSL"
    ELSE
           Texture = Texture & ":" & ""
    END IF
'Loam
  Case Clay >= 7.0 And Clay < 27.0 And Silt >= 28.0 And Silt < 50.0 And Sand <= 52.0

Texture = Texture & "CASE Clay >= 7.0 And Clay < 27.0 And Silt >= 28.0 And Silt < 50.0 And Sand <= 52.0" & vbcrlf

       Texture = Texture & ":" & "L"

  CASE Silt >= 50.0 AND Clay >= 12.0 AND Clay < 27.0

Texture = Texture & "CASE Silt >= 50.0 AND Clay >= 12.0 AND Clay < 27.0" & vbcrlf

       Texture = Texture & ":" & "SiL"

  CASE Silt >= 50.0 AND Silt < 80.0 AND Clay < 12.0

Texture = Texture & "CASE Silt >= 50.0 AND Silt < 80.0 AND Clay < 12.0" & vbcrlf

       Texture = Texture & ":" & "SiL"

  CASE Silt >= 80.0 AND Clay < 12.0

Texture = Texture & "CASE Silt >= 80.0 AND Clay < 12.0" & vbcrlf

       Texture = Texture & ":" & "Si"

  CASE Clay >= 20.0 AND Clay < 35.0 AND Silt < 28.0 AND Sand > 45.0

Texture = Texture & "CASE Clay >= 20.0 AND Clay < 35.0 AND Silt < 28.0 AND Sand > 45.0" & vbcrlf

       Texture = Texture & ":" & "SCL"

  CASE Clay >= 27.0 AND Clay < 40.0 AND Sand > 20.0 AND Sand <= 45.0

Texture = Texture & "CASE Clay >= 27.0 AND Clay < 40.0 AND Sand > 20.0 AND Sand <= 45.0" & vbcrlf

        Texture = Texture & ":" & "CL"

  CASE Clay >= 27.0 AND Clay < 40.0 AND Sand <= 20.0

Texture = Texture & "CASE Clay >= 27.0 AND Clay < 40.0 AND Sand <= 20.0" & vbcrlf

        Texture = Texture & ":" & "SiCL"

  CASE Clay >= 35.0 AND Sand > 45.0

Texture = Texture & "CASE Clay >= 35.0 AND Sand > 45.0" & vbcrlf

        Texture = Texture & ":" & "SC"

  CASE Clay >= 40.0 AND Silt >= 40.0

Texture = Texture & "CASE Clay >= 40.0 AND Silt >= 40.0" & vbcrlf

        Texture = Texture & ":" & "SiC"

  CASE Clay >= 40.0 AND Silt < 40.0 AND Sand <= 45.0

Texture = Texture & "CASE Clay >= 40.0 AND Silt < 40.0 AND Sand <= 45.0" & vbcrlf

        Texture = Texture & ":" & "C"

 END SELECT

 OUTPUT = LCase(Texture)
END FUNCTION

'-----------------------------------------------------------------------------------------------------------------------------------------------

Function AsymArithRnd( x,  Factor )

'Asymmetric arithmetic rounding - rounds .5 up always
'factor is multiple of 10
If Factor = 0 Then Factor = 1  'added this for whole number notation within the
                               'analyte format field .0

AsymArithRnd = Int(CCur(x * Factor + 0.5)) / Factor

End Function

