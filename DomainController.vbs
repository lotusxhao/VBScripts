Set objDomain = GetObject("LDAP://RootDSE")
objDC = objDomain.Get("dnsHostName")
Wscript.Echo objDC