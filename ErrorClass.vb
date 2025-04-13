Public Class ErrorClass

    Protected m_ErrNum As Integer
    Protected m_ErrDsc As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the Last Error number set by the inheriting class    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ErrNum() As Integer
        Get
            ErrNum = m_ErrNum
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the last Error Message set by the inheriting class    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ErrDsc() As String
        Get
            ErrDsc = m_ErrDsc
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clears the last Error seet by the inheriting class     ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[U910KRZI]	05/05/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ErrClr()
        m_ErrDsc = ""
        m_ErrNum = 0
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
