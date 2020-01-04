#pragma checksum "C:\Users\Durgesh\source\repos\FinResearch\Views\Home\Index.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "738a68ab71ad7ddbb6a0d09db2174a7079361ec4"
// <auto-generated/>
#pragma warning disable 1591
[assembly: global::Microsoft.AspNetCore.Razor.Hosting.RazorCompiledItemAttribute(typeof(AspNetCore.Views_Home_Index), @"mvc.1.0.view", @"/Views/Home/Index.cshtml")]
namespace AspNetCore
{
    #line hidden
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.Mvc.Rendering;
    using Microsoft.AspNetCore.Mvc.ViewFeatures;
#nullable restore
#line 1 "C:\Users\Durgesh\source\repos\FinResearch\Views\_ViewImports.cshtml"
using FinResearch;

#line default
#line hidden
#nullable disable
#nullable restore
#line 2 "C:\Users\Durgesh\source\repos\FinResearch\Views\_ViewImports.cshtml"
using FinResearch.Models;

#line default
#line hidden
#nullable disable
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"738a68ab71ad7ddbb6a0d09db2174a7079361ec4", @"/Views/Home/Index.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"3faaf773cce3914737b3a76ef26edbb3a3c28549", @"/Views/_ViewImports.cshtml")]
    public class Views_Home_Index : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<dynamic>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
#nullable restore
#line 1 "C:\Users\Durgesh\source\repos\FinResearch\Views\Home\Index.cshtml"
  
	ViewData["Title"] = "Home Page";

#line default
#line hidden
#nullable disable
            WriteLiteral(@"

<div id=""my""></div>


<script src=""https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js""></script>
<script src=""https://bossanova.uk/jexcel/v3/jexcel.js""></script>
<link rel=""stylesheet"" href=""https://bossanova.uk/jexcel/v3/jexcel.css"" type=""text/css"" />

<script src=""https://bossanova.uk/jsuites/v2/jsuites.js""></script>
<link rel=""stylesheet"" href=""https://bossanova.uk/jsuites/v2/jsuites.css"" type=""text/css"" />

<script>
	var serverdata = [];
	$.ajax({
		type: ""GET"",
		url: ""https://localhost:44338/FinanceStatements/GetStatements"",
		dataType: ""json"",
		async: false,
		contentType: 'application/json; charset=utf-8',
		success: function (data) {
			serverdata = JSON.parse(data);
		},
		error: function () {
			alert(""Error occured!!"")
		}
	});


	var resArr = serverdata.map(function (x) {
		return Object.keys(x);
	});
	var longest = resArr.reduce((a, b) => (a.length > b.length ? a : b), []);
	var valuearray = serverdata.map(function (x) {
		return Object.values(");
            WriteLiteral(@"x);
	});
	var catIndex = longest.indexOf('CATEGORY', 0);
	//var sortestArr = valuearray;

	//var p = valuearray.map(function (x) {
		
	//	if (x.length == 1) {
			
	//		var item = x[0];
	//		var i;
	//		for ( i = 0; i < longest.length; i++) {
	//			if (i != catIndex) {
	//				x[i]= '-';
	//			}
	//			else {
	//				x[i]= item;
	//			}
				
	//		}
	//		console.log(x);
			
			
	//	}
	//	return x;
	//});
	//valuearray = valuearray.filter(x => x.length > 1);
	//sortestArr = sortestArr.filter(x => x.length == 1);
	//sortestArr.forEach(function (element) {

	//	for ( var i = 0; i < longest.length; i++) {
	//		element.push('-');
	//	}
	//});
	//valuearray.push(sortestArr);
	//console.log(valuearray);
	//console.log(sortestArr);
	debugger;
	function numberToLetters(nNum) {
    var result;
    if (nNum <= 26) {
        result = letter(nNum);
    } else {
        var modulo = nNum % 26;
        var quotient = Math.floor(nNum / 26);
        if (modulo === 0) {
            resu");
            WriteLiteral(@"lt = letter(quotient - 1) + letter(26);
        } else {
            result = letter(quotient) + letter(modulo);
        }
    }

    return result;
	}
	var arrNested = [];
	for (i = 1; i <= longest.length; i++) {
		arrNested.push({""title"":numberToLetters(i) });
	}

	console.log(arrNested);
	var mytable=$('#my').jexcel({
		
    minDimensions:[13,13],
  
		data: valuearray,
		colHeaders: longest,
		nestedHeaders: [arrNested]


	});
	


function letter(nNum) {
    var a = ""A"".charCodeAt(0);
    return String.fromCharCode(a + nNum - 1);
}
	
	console.log(catIndex);
	mytable.moveColumn(catIndex, 0);
	mytable.setWidth(0, 400);
	for ( var i = 1; i < longest.length; i++) {
			mytable.setWidth(i, 100);
		}

</script>");
        }
        #pragma warning restore 1998
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.ViewFeatures.IModelExpressionProvider ModelExpressionProvider { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IUrlHelper Url { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IViewComponentHelper Component { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IJsonHelper Json { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IHtmlHelper<dynamic> Html { get; private set; }
    }
}
#pragma warning restore 1591
