<div align="center">

## Caption Animation \- Timer Control


</div>

### Description

Beginner level discussion on proper use of the Timer Control. Give your About Box a bit of sex appeal.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ross Ylitalo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ross-ylitalo.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ross-ylitalo-caption-animation-timer-control__1-37114/archive/master.zip)





### Source Code

```
Private m_strFullCaption As String
Private m_strCurrentCaption As String
Private m_ynLeftRight As Boolean
Private Sub Form_Load()
  m_strFullCaption = Me.Caption ' Don't mutate m_strFullCaption after this!
  Me.Caption = ""
  m_strCurrentCaption = ""
  m_ynLeftRight = True
End Sub
Private Sub tmr_Timer()
  If m_ynLeftRight Then
    ' Show caption from left to right.
    If Len(m_strCurrentCaption) < Len(m_strFullCaption) Then
      m_strCurrentCaption = Left(m_strFullCaption, Len(m_strCurrentCaption) + 1)
      Me.Caption = m_strCurrentCaption
    Else
      m_ynLeftRight = Not m_ynLeftRight
    End If
  Else
    ' Scroll caption off of screen, scrolling Left.
    If Len(m_strCurrentCaption) > 0 Then
      m_strCurrentCaption = Right(m_strFullCaption, Len(m_strCurrentCaption) - 1)
      Me.Caption = m_strCurrentCaption
    Else
      m_ynLeftRight = Not m_ynLeftRight
    End If
  End If
End Sub
```

