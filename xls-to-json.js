var input="biography of sutra 3 for json";  // input folder name

/////////////////////////////////////////////////////
var fs=require("fs");
var xls_json=require("xls-to-json");
var mkdirp=require("mkdirp");
var path=require("path").dirname;
var list=fs.readdirSync(input);

var xlstojson=function(xls) {
	var name=xls.split(".");
	var outfn="out/"+name[0];
	var dir=path(outfn);
	mkdirp.sync(dir);
	xls_json({
	  input: __dirname + "/"+input+"/"+name[0]+".xls",
	  output: __dirname + "/out/"+name[0]+".json"
	}, function(err, result) {
	  
	  if(err) {
	    console.error(err);
	  } 
	});
} 

list.map(xlstojson);