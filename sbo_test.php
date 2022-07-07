<html>
 <head>
  <title>Teste Conexão DI API</title>
 </head>
 <body>
  <?php //Utilize a tag absoluta

  $oComp = new COM("SAPbobsCOM.Company") or die("No connection");
  $oComp->Server = "hostname";
  $oComp->LicenseServer = "hostname";
  $oComp->DbUserName = "userdb"; //Usuário do banco de dados
  $oComp->DbPassword = "user@db"; //Senha do banco de dados
  $oComp->DBServerType = 11; //Tipo 7 indica o SGBD MSSQL 2012
  $oComp->UseTrusted = false; //False = sem autenticação com Windows
  $oComp->UserName = "usersap"; //Usuário do SAP
  $oComp->Password = "user@sap"; //Senha do usuário do SAP
  $oComp->CompanyDB = "SBO_DEMO_PRD"; //Seu banco de dados

  //Testaremos sua conexão – se conectado
  try {
      echo $oComp->Connect;
      echo "<br><br>";
      $oComp->StartTransaction();
      echo $oComp->CompanyName . ‘<br \>’;
      $oComp->Disconnect;

  //Senão, mensagem de erro
  } catch (com_exception $expt) {
      echo $expt->getMessage();
      echo "<br><br>" . $oComp->GetLastErrorDescription;
  }

  ?>
 </body>
</html>