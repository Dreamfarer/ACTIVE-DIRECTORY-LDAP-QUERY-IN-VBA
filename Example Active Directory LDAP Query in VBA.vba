Function ACTIVEDIRECTORY(sUser, sPassword) As Variant

    'Declaring variables
    Dim oDS, ooAuth, oRecordSet
    
    'Set destinguished name (DN) and LDAP root
    sDN = "cn=" & sUser & ",dc=IT Accounts,dc=intranet,dc=example,dc=ch" 'Look up the destinguished name (DN) of your account with which you are trying to login to the active directory (AD).
    sRoot = "LDAP://intranet.example.ch/dc=intranet,dc=example,dc=ch" 'Domain of server hosting the active directory (AD).
    
    'Connect to active directory (AD) server
    Set oDS = GetObject("LDAP:")
    Set oAuth = oDS.OpenDSObject(sRoot, sDN, sPassword, &H200) 'We use "OpenDSObject" because we want to access the active directory (AD) with a different user.
    
    'Create connection
    Dim oConn: Set oConn = CreateObject("ADODB.Connection")
    oConn.Provider = "ADSDSOObject"
    oConn.Open "Ads Provider", sDN, sPassword
    
    'Build query
    sBase = "<" & sRoot & ">;"
    sFilter = "(&(objectCategory=person)(objectClass=user)(memberOf=cn=group1,dc=intranet,dc=example,dc=ch));" 'Filter to only get objects that match this criteria. SQL equivalent: WHERE group="group1". "memberOf" is used to get all object that match this destinguished name (DN).
    sAttributes = "cn,mail,telephoneNumber,mobile,physicalDeliveryOfficeName,title;" 'Objects you want to retrieve. SQL equivalent: SELECT mail
    sScope = "subtree"
    sLDAPQuery = sBase & sFilter & sAttributes & sScope
    
    'Execute Query
    Set oRecordSet = oConn.Execute(sLDAPQuery) 'Record set containing the retrieved object
    
    'Declare Array
    Dim aQueryCapture() As Variant
    ReDim aQueryCapture(oRecordSet.RecordCount - 1, 5) 'Resize array to match the record set length
    
    'Fill Array (If null, replace it with "")
    Dim counter As Integer
    counter = 0
    While Not oRecordSet.EOF
        If Not IsNull(oRecordSet("cn").Value) Then
            aQueryCapture(counter, 0) = oRecordSet("cn").Value
        Else
            aQueryCapture(counter, 0) = ""
        End If
        
        If Not IsNull(oRecordSet("title").Value) Then
            aQueryCapture(counter, 1) = oRecordSet("title").Value
        Else
            aQueryCapture(counter, 1) = ""
        End If
        
        If Not IsNull(oRecordSet("mail").Value) Then
            aQueryCapture(counter, 2) = oRecordSet("mail").Value
        Else
            aQueryCapture(counter, 2) = ""
        End If
        
        If Not IsNull(oRecordSet("telephoneNumber").Value) Then
            aQueryCapture(counter, 3) = oRecordSet("telephoneNumber").Value
        Else
            aQueryCapture(counter, 3) = ""
        End If
        
        If Not IsNull(oRecordSet("mobile").Value) Then
            aQueryCapture(counter, 4) = oRecordSet("mobile").Value
        Else
            aQueryCapture(counter, 4) = ""
        End If
        
        If Not IsNull(oRecordSet("physicalDeliveryOfficeName").Value) Then
            aQueryCapture(counter, 5) = oRecordSet("physicalDeliveryOfficeName").Value
        Else
            aQueryCapture(counter, 5) = ""
        End If

        counter = counter + 1
        oRecordSet.MoveNext 'Go to next entry in record set
    Wend
    
    ACTIVEDIRECTORY = aQueryCapture 'Return array
    
End Function