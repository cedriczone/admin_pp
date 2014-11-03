<!--#include file="verif_ident.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">&nbsp;
  <p><a href="dispos_gav.asp">Gestion des disponibilit&eacute;s intervenants</a></p>
  <p><a href="gene_gav_coord.asp">G&eacute;n&eacute;rer le planning coordinateurs</a></p>
  <p><form id="form2" name="form2" method="post" action="remplacement.asp">
  <p>Echanger le planning de 2 avocats :</p>
  <p>Date (jj/mm/aaaa) : 
    <input name="date_echange" type="text" id="date_echange" size="12" />
    <input type="submit" name="button2" id="button2" value="Visualiser" />
  </p>
</form></p>
  <p><a href="gene_planning_gav.asp">G&eacute;n&eacute;ration du planning</a></p>
  <p><a href="compteurs_gav.asp">Compteurs de d&eacute;signations GAV</a></p>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
