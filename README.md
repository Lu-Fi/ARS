# Access Request System

First I would like to explain some of my motivations for the development as well as views regarding secure administration in Active Directory environments.

---

In my opinion there are only 2 really useful ways to increase security in Active Directory Domains.

1. all administrative operations are performed exclusively by secure systems (PAW's). This includes not only administrative activities on domain controllers but all administrative activities divided by tiers.

As an example: Tier0 = Domain Controller, Tier1 = Member Systems, Tier2 = Clients.

Such an approach is unfortunately very difficult to implement and can be very expensive depending on the environment.

2. the range and duration of administrative privileges is limited to the most necessary.
Privileges are requested for a certain duration and expire afterwards.

---

Because of the costs and benefits, I looked more closely at the second option and what options are available. I noticed that Microsoft has already implemented a way to assign group memberships time-bound in Server 2016. Here a special LDAP extension is used to specify the number of seconds how long a membership exists (LDAP_SERVER_LINK_TTL extended control (OID = 1.2.840.113556.1.4.2309)).

Unfortunately, the assignment of such memberships is currently only possible with Powershell, which makes the whole thing a bit uncomfortable.

This is where my tool comes in. I provide a webinterface, with which it is possible to request and remove group memberships time-bound. In addition, it can be specified that the administrative accounts, which are managed in the application, e.g. after the expiration of the last group membership, are deactivated and the password resets automatically. These accounts are then reactivated when a new membership is requested and a new password can be retrieved.

Of course, not just anyone can request any group membership. The application aims to ensure that each administrator has at least one Office and one or more Administrative accounts.

The web application is used with the Office account. To improve security, the login requires an Active Directory on the one hand and an Azure login for the Office Account (This is also how MFA can be implemented).

The Office account is then assigned the administrative accounts within the application, which the user is allowed to manage. The assignment of the accounts takes place either in the application or via an extension attribute in the Active Directory.

The administrative accounts are then either assigned to so-called "request groups", via these groups it is then defined which group memberships may be requested (managedBy attribute) or a direct assignment is made within the application (SQL database).

The administrative process then looks as follows:
1. the office user logs on to the application (incl. MFA)
2. he chooses an administrative account in the web interface and requests a group membership for x hours.
3 [Optional] A new password is requested.
4. the administrative activities are executed
5. [Optional] The time span is extended
6. after the time period expires, Active Directory removes the group memberships by itself, the Kerberos tokens are invalid.
7. [Optional] The application disables the Administrative account and sets a random password.

I have been able to analyze attacks in the past, and the procedure is almost always the same. First, threat actors try to compromise an account via social engineering, fake company websites, mail, etc. It does not matter which rights this account has. Next, the environment is analyzed. Such an analysis quickly reveals possible configuration errors on systems as well as attack vectors on services and applications. If an attack target is not already found here (e.g. the user has local administration rights on his system or on a server and can thus compromise the local SAM), the account is misused to compromise other systems via mails and documents, if necessary.

Actually, at this point you can already assume that if 0,1,2 administrative rights have been obtained in a tier, the complete tier is more or less lost.

---

![screen_1](https://user-images.githubusercontent.com/1177251/150324132-6d614dbb-0759-4995-97ee-fd7adbb1fc1a.png)
![screen_2](https://user-images.githubusercontent.com/1177251/150324137-461cc29b-b9bf-4859-b256-8d17a2ccd4f5.png)
![screen_3](https://user-images.githubusercontent.com/1177251/150324139-d9572ce4-4ac5-4762-8272-5c90ab98f8b1.png)
![screen_4](https://user-images.githubusercontent.com/1177251/150324143-545afd3b-8fa4-4894-930f-fd0643120d57.png)


# Autostart with IIS

c:\windows\system32\inetsrv\config\applicationHost.config

    <applicationPools>
        <add name="ARS" autoStart="true" managedRuntimeVersion="v4.0" startMode="AlwaysRunning">
            <processModel identityType="SpecificUser" userName="DOMAIN\USER" password="" loadUserProfile="false" idleTimeout="00:00:00" maxProcesses="0" />
            <recycling>
                <periodicRestart time="00:00:00">
                    <schedule>
                        <clear />
                    </schedule>
                </periodicRestart>
            </recycling>
        </add>
    </applicationPools>

    <sites>
        <site name="ARS" id="2" serverAutoStart="true">
            <application path="/" applicationPool="ARS" preloadEnabled="true">
                <virtualDirectory path="/" physicalPath="C:\inetpub\wwwroot\ARS" />
            </application>
            <bindings>
                <binding protocol="https" bindingInformation="*:443:ars.domain.net" sslFlags="0" />
            </bindings>
        </site>
    </sites>

    <location path="ARS">
        <system.webServer>
            <security>
                <authentication>
                    <windowsAuthentication enabled="true" useKernelMode="false">
                        <extendedProtection tokenChecking="None" />
                        <providers>
                            <clear />
                            <add value="Negotiate:Kerberos" />
                        </providers>
                    </windowsAuthentication>
                    <anonymousAuthentication enabled="false" />
                </authentication>
                <access sslFlags="Ssl" />
            </security>
        </system.webServer>
    </location>


https://docs.microsoft.com/en-us/iis/get-started/whats-new-in-iis-8/iis-80-application-initialization


# SQL Scheme
(Create empty DB called "**Ars**" first)

    USE [Ars]
    GO

    SET ANSI_NULLS ON
    GO
    SET QUOTED_IDENTIFIER ON
    GO
    CREATE TABLE [dbo].[ars_assignments](
      [aID] [uniqueidentifier] NOT NULL,
      [uSID] [varchar](256) NOT NULL,
      [aSID] [varchar](256) NOT NULL,
      [flag] [int] NOT NULL
    ) ON [PRIMARY]
    GO

    SET ANSI_NULLS ON
    GO
    SET QUOTED_IDENTIFIER ON
    GO
    CREATE TABLE [dbo].[ars_audit](
      [aSID] [varchar](256) NOT NULL,
      [aName] [varchar](256) NOT NULL,
      [uSource] [varchar](128) NULL,
      [uName] [varchar](256) NOT NULL,
      [action] [varchar](256) NOT NULL,
      [result] [int] NOT NULL,
      [ts] [datetime] NULL
    ) ON [PRIMARY]
    GO

    SET ANSI_NULLS ON
    GO
    SET QUOTED_IDENTIFIER ON
    GO
    CREATE TABLE [dbo].[ars_user](
      [uSID] [varchar](256) NOT NULL,
      [pwdLastSet] [bigint] NULL
    ) ON [PRIMARY]
    GO

    SET ANSI_NULLS ON
    GO
    SET QUOTED_IDENTIFIER ON
    GO
    CREATE TABLE [dbo].[ars_users](
      [uSID] [varchar](256) NOT NULL,
      [flag] [int] NOT NULL
    ) ON [PRIMARY]
    GO
    ALTER TABLE [dbo].[ars_assignments] ADD  CONSTRAINT [DF_ars_assignments_aID]  DEFAULT (newid()) FOR [aID]
    GO
    ALTER TABLE [dbo].[ars_users] ADD  CONSTRAINT [DF_ars_users_flag]  DEFAULT ((0)) FOR [flag]
    GO

