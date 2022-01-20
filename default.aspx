<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="default.aspx.vb" Inherits="Ars._default" %>

<!DOCTYPE html>
<html lang="de">
<head runat="Server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta charset="utf-8" />
    <title>ARS</title>
    <link href="~/favicon.ico" rel="shortcut icon" type="image/x-icon" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, shrink-to-fit=no">
    <script src="inc/js/jquery/jquery-3.5.1.min.js"></script>
    <script src="inc/js/jqueryui/jquery-ui.min.js"></script>
    <script src="inc/js/moment/moment.js"></script>
    <script src="inc/js/bootstrap/bootstrap.min.js"></script>
    <script src="inc/js/datatables/jquery.dataTables.min.js"></script>
    <script src="inc/js/datatables/dataTables.bootstrap5.min.js"></script>
    <script src="inc/js/datatables/datetime.js"></script>
    <script src="inc/js/datatables/pageresize.js"></script>
    <script src="inc/js/typeahead/bootstrap-typeahead.js"></script>
    <script src="inc/js/ars.js"></script>
    <link href="inc/css/bootstrap/bootstrap.min.css" rel="stylesheet" />
    <link href="inc/css/bootstrap/bootstrap-icons.css" rel="stylesheet" />
    <link href="inc/css/datatables/dataTables.bootstrap5.min.css" rel="stylesheet" />
    <link href="inc/css/ars.css" rel="stylesheet" />
</head>
<body>
    <header class="navbar navbar-dark sticky-top bg-dark flex-md-nowrap p-0 shadow">
      <a class="navbar-brand col-md-3 col-lg-2 me-0 px-3" href="#">ARS - made time bound security easy</a>
      <ul class="navbar-nav">
        <li class="nav-item text-nowrap rm-5" title="Active Directory Authenticated" runat="server" id="arsAdAuth" visible="false">
            <button type="button" class="btn btn-primary disabled btn-sm btn-ars-state">AD</button>
        </li>
        <li class="nav-item text-nowrap rm-5" title="Azure Authenticated" runat="server" id="arsAzAuth" visible="false">
            <button type="button" class="btn btn-primary disabled btn-sm btn-ars-state" id="btnAzAuth" runat="server"><span>AZURE</span><span></span></button>
        </li>
        <li class="nav-item text-nowrap">
            <form class="navbar-form" action="./default.aspx" method="POST" id="authForm">
                <input type="hidden" id="ActionId" Value="" runat="server" />
                <input type="button" id="btnAction" Value="" class="btn btn-dark" runat="server" />
            </form>
        </li>
        <li>
          <button class="navbar-toggler position-relative d-md-none collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#arsNav" aria-controls="arsNav" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
          </button>
        </li>
      </ul>
    </header>

    <div class="container-fluid" id="arsContainer" runat="server">
      <div class="row">
        <nav class="col-md-3 col-lg-2 d-md-block bg-light sidebar collapse" id="arsNav" runat="server" visible="false" >
          
          <div id="arsNavPersonalMenu" class="position-sticky pt-3">
            <h3>Accounts</h3>
            <ul class="nav flex-column">
              <li class="nav-item"><a class="nav-link active" href="#" data-menu="groups">Permissions</a></li>
              <li class="nav-item"><a class="nav-link" href="#" data-menu="advanced">Advanced Features</a></li>
              <li class="nav-item"><a class="nav-link" href="#" data-menu="audit">Audit Log</a></li>
            </ul>
          </div>

          <div id="arsNavAdminMenu" class="position-sticky pt-3" runat="server" visible="false" >
            <h3>Admin</h3>
            <ul class="nav flex-column">
              <li class="nav-item"><a class="nav-link" href="#" data-menu="adm_accounts">Accounts</a></li>
              <li class="nav-item"><a class="nav-link" href="#" data-menu="adm_assignments">Assignments</a></li>
                <li class="nav-item"><a class="nav-link" href="#" data-menu="adm_audit">Admin Audit Log</a></li>
            </ul>
          </div>
        </nav>

        <main class="col-md-9 offset-md-2 col-lg-10 px-md-4 position-fixed" id="arsMain" runat="server" visible="false" >
         
          <div id="arsBusy">
            <div class="spinner"></div>
            <br/>
            Loading...
          </div>
          <div id="arsCw" class="arsContentWrapper">

          </div>
        </main>

      </div>

      <div id="arsMessage_AccessDenied" class="arsMessage" runat="server" visible="false">
        <h1>Access denied</h1>
        <hr />
        <div>You are not permitted to use this System.</div>
      </div>

      <div id="arsMessage_Login" class="arsMessage" runat="server" visible="false">
        <h1>Please login</h1>
        <hr />
        <div>In order to use this system you must login.</div>
      </div>

      <div id="arsMessage_AuthError" class="arsMessage" runat="server" visible="false">
        <h1>Authentication failure</h1>
        <hr />
        <div id="arsMessage_AuthErrorMessage" runat="server"></div>
      </div>

    </div>
    <div class="modal fade" id="arsModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"></h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">

                </div>
                <div class="modal-footer">
                    
                </div>
            </div>
        </div>
    </div>
</body>
</html>

