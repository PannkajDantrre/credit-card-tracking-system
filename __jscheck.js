var fso = new ActiveXObject('Scripting.FileSystemObject');
var file = fso.OpenTextFile('E:\\Credit Card Tracking System\\app.js', 1);
var source = file.ReadAll();
file.Close();
try {
  new Function(source);
  WScript.Echo('JS_OK');
} catch (e) {
  WScript.Echo('JS_ERROR:' + e.message);
}
