function myFunction() {
var string = 'あいう,えお?かき';
var searchKeyword = /[,?]/g;
var result;

var end = searchKeyword.length;

//while (result = searchKeyword.exec(string)) { 
//  console.log(result.index);
 
//}


var output = string.replace(searchKeyword, searchKeyword+"\n")
console.log(output);
}