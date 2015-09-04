var input="biography of sutra 3";  // input folder name

/////////////////////////////////////////////////////
var fs=require("fs");
var xlsx_json=require("xlsx-to-json");
var mkdirp=require("mkdirp");
var path=require("path").dirname;
var list=fs.readdirSync(input);

var xlsxtojson=function(xlsx) {
	var name=xlsx.split(".");
	var outfn="out/"+name[0];
	var dir=path(outfn);
	mkdirp.sync(dir);
	xlsx_json({
	  input: __dirname + "/"+input+"/"+name[0]+".xlsx",
	  output: __dirname + "/out/"+name[0]+".json"
	}, function(err, result) {
	  
	  if(err) {
	    console.error(err);
	  } 
	});
} 
list.map(xlsxtojson);
