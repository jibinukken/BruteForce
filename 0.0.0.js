WSH.echo ( "Run Winzip" );

var pwchars = "01233456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
var b = 0;
var n, i ;

WSH.echo ( pwchars.length );

//var name=prompt("your name here","")
//WSH.echo ( name + "!" );

var i=0;
 for (n=1;n<=1;n++) /*take n to be 32.*/
 {	
   for (i=0;i<=pwchars.length-1;i++)
   {
    var pwchar = new Array();
    
    var pw = pwchars.substr(i,n);
    //var shell = WScript.CreateObject("WScript.shell");
    var shell = new ActiveXObject("WScript.shell");
    shell.run( "winzip32 -e -s" + pw + " " + "example.zip \"C:\Zipextract\"")
    shell.run( "taskkill.exe /im winzip32.exe" )
    /*if 
    ()
     exists
     exit;
     0xc0000142*/
shell.run( "taskkill.exe /f /im winzip32.exe" )
   }
 }
