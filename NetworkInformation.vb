Imports System.Net.NetworkInformation
Module NetWorkInformation
    Sub GetNetworkInformation()
        xtrace("______________________________________________________________________")
        xtrace("")
        Dim HostName1 As String = Environment.MachineName
        xtrace("Host Name ..................... : " & HostName1, 1)

        Dim HostName2 As String = IPGlobalProperties.GetIPGlobalProperties().HostName
        xtrace("Host Name 2 ................... : " & HostName2)

        Dim HostName3 As String = System.Net.Dns.GetHostName()
        xtrace("Host Name 3 ................... : " & HostName3)

        Dim UserName1 As String = Environment.UserName
        xtrace("User name ..................... : " & UserName1, 1)

        Dim UserName2 As String = System.Security.Principal.WindowsIdentity.GetCurrent().Name
        xtrace("User name 2 ................... : " & UserName2)

        Dim UserSID As String = System.Security.Principal.WindowsIdentity.GetCurrent.User.ToString
        xtrace("User SID ...................... : " & UserSID, 1)

        '---- IPProp 
        Dim UserDomainName As String = Environment.UserDomainName
        xtrace("UserDomainName ................ : " & UserDomainName, 1)

        Dim DomainName As String = IPGlobalProperties.GetIPGlobalProperties().DomainName
        If DomainName.Length < 2 Then
            xtrace("Fully Qualified Domain Name ... : N.a.", 1)
        Else
            xtrace("Fully Qualified Domain Name ... : " & DomainName, 1)
        End If

        Dim strIPAddress As String = System.Net.Dns.GetHostEntry(HostName3).AddressList(0).ToString()
        xtrace("IPAddress ..................... : " & strIPAddress, 1)

        '---- Network
        Dim IPPropperties As IPGlobalProperties
        IPPropperties = IPGlobalProperties.GetIPGlobalProperties

        Dim Nics() As NetworkInterface
        Dim Nic As NetworkInterface
        Nics = NetworkInterface.GetAllNetworkInterfaces
        Dim NicCnt = Nics.Length

        If NicCnt > 0 Then
            xtrace("Nr of NIC interfaces .......... : " & NicCnt.ToString, 1)
            Dim NicNr As Integer = 0

            For Each Nic In Nics
                NicNr += 1
                Dim NicProp As IPInterfaceProperties
                NicProp = Nic.GetIPProperties
                xtrace(" ", 1)
                Dim NicName As String = Nic.Name
                If Left(NicName, 8) = "Ethernet" Then
                    xtrace("Ethernet adapter " & NicName & "  (" & NicNr.ToString & ")", 1)
                Else
                    xtrace(NicName & "  (" & NicNr.ToString & ")", 1)
                End If

                Dim IsLoopBack As Boolean = (Nic.NetworkInterfaceType = NetworkInterfaceType.Loopback)

                If Not IsLoopBack Then xtrace("  Connection specific DNS suffix ...... : " & NicProp.DnsSuffix, 1)

                Dim NicDescr As String = Nic.Description
                xtrace("  Description ......................... : " & NicDescr, 1)

                Dim MAC1 = Nic.GetPhysicalAddress().ToString()
                Dim MAC2 = Left(MAC1, 2) & "-" & Mid(MAC1, 3, 2) & "-" & Mid(MAC1, 5, 2) & "-" & Mid(MAC1, 7, 2) & "-" & Mid(MAC1, 9, 2) & "-" & Mid(MAC1, 11, 2)
                'xtrace("  Physical Address (MAC) .............. : " & MAC1)
                xtrace("  Physical Address (MAC) .............. : " & MAC2, 1)

                xtrace("  Interface type ...................... : " & Nic.NetworkInterfaceType.ToString, 1)   ' ToString translates the integer to a meaningfull string
                xtrace("  Operational status .................. : " & Nic.OperationalStatus.ToString(), 1)

                Dim Versions As String = ""
                If (Nic.Supports(NetworkInterfaceComponent.IPv4)) Then
                    Versions = "IPv4"
                End If
                If (Nic.Supports(NetworkInterfaceComponent.IPv6)) Then
                    If Versions.Length > 2 Then
                        Versions = Versions & ", IPv6"
                    Else
                        Versions = "IPv6"
                    End If
                End If
                xtrace("  IP version .......................... : " & Versions, 1)

                If IsLoopBack Then Continue For

                If Nic.Supports(NetworkInterfaceComponent.IPv4) Then
                    Dim IPV4Prop As IPv4InterfaceProperties
                    IPV4Prop = NicProp.GetIPv4Properties
                    xtrace("  MTU.................................. : " & IPV4Prop.Mtu, 1)

                    If IPV4Prop.UsesWins Then
                        Dim winsServers As IPAddressCollection
                        winsServers = NicProp.WinsServersAddresses
                        ShowAddressList(winsServers, "WINS Server .........................")
                    Else
                        xtrace("  WINS Servers ........................ : N.a.")
                    End If
                End If

                ShowAddressList(NicProp.DnsAddresses, "DNS Server ..........................")

                For Each Gateway As GatewayIPAddressInformation In NicProp.GatewayAddresses
                    xtrace("  Gateway Address ..................... : " & Gateway.Address.ToString, 1)
                Next

                xtrace("  DNS enabled ......................... : " & NicProp.IsDnsEnabled, 1)
                xtrace("  Dynamically configured DNS .......... : " & NicProp.IsDynamicDnsEnabled, 1)
                xtrace("  Receive Only ........................ : " & Nic.IsReceiveOnly, 1)
                xtrace("  Multicast ........................... : " & Nic.SupportsMulticast, 1)
            Next
        End If

    End Sub

    Sub ShowAddressList(AddrList As IPAddressCollection, Label As String)
        If AddrList.Count = 0 Then
            xtrace("  " & Label & " : None")
        Else
            Dim IPAddr
            For Each IPAddr In AddrList
                xtrace("  " & Label & " : " & IPAddr.ToString, 1)
            Next
        End If
    End Sub

End Module
