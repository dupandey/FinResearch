#pragma checksum "C:\FinResearchWeb\Views\Home\Index.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "67e2fdb3e93a9b28ca548d8dca66769c46b55322"
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
#line 1 "C:\FinResearchWeb\Views\_ViewImports.cshtml"
using FinResearch;

#line default
#line hidden
#nullable disable
#nullable restore
#line 2 "C:\FinResearchWeb\Views\_ViewImports.cshtml"
using FinResearch.Models;

#line default
#line hidden
#nullable disable
#nullable restore
#line 3 "C:\FinResearchWeb\Views\_ViewImports.cshtml"
using FinResearch.Models.AccountViewModels;

#line default
#line hidden
#nullable disable
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"67e2fdb3e93a9b28ca548d8dca66769c46b55322", @"/Views/Home/Index.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"d7c65629757e2741f005f142801ac97cbd0e9213", @"/Views/_ViewImports.cshtml")]
    public class Views_Home_Index : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<dynamic>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
#nullable restore
#line 1 "C:\FinResearchWeb\Views\Home\Index.cshtml"
  
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
	var winWidth = $(window).width();
	var serverdata=[];
	$.ajax({
		type: ""GET"",
		url: ""/FinanceStatements/GetStatements"",
		dataType: ""json"",
		async: false,
		contentType: 'application/json; charset=utf-8',
		success: function (data) {
			serverdata = data;
		},
		error: function () {
			alert(""Error occured!!"")
		}
	});
	var themedata = [];
	$.ajax({
		type: ""GET"",
		url: ""/FinanceStatements/Gettheme"",
		dataType: ""json"",
		async: false,
		contentType: 'application/json; charset=utf-8',
		success: function (data) {
		themedata = d");
            WriteLiteral(@"ata;
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
		return Object.values(x);
	});
	var catIndex = longest.indexOf('LineItem', 0);
	//var sortestArr = valuearray;
	console.log(longest);
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
	//co");
            WriteLiteral(@"nsole.log(valuearray);
	//console.log(sortestArr);

	function numberToLetters(nNum) {
		var result;
		if (nNum <= 26) {
			result = letter(nNum);
		} else {
			var modulo = nNum % 26;
			var quotient = Math.floor(nNum / 26);
			if (modulo === 0) {
				result = letter(quotient - 1) + letter(26);
			} else {
				result = letter(quotient) + letter(modulo);
			}
		}

		return result;
	}
	var arrNested = [];
	for (i = 1; i <= longest.length; i++) {
		arrNested.push({ ""title"": numberToLetters(i) });
	}

	var filterdata = null;
	var mytable = $('#my').jexcel({
		csvHeaders: true,
		minDimensions: [5, 13],
		tableOverflow: true,
		tableWidth: (winWidth - 40) + 'px',
		tableHeight: ""800px"",
		freezeColumns: 1,
		wordWrap: true,
		data: valuearray,
		colHeaders: longest,
		nestedHeaders: [arrNested],
		updateTable: function (instance, cell, col, row, val, label, cellName) {
			var header = $(""thead tr td[data-x='"" + col + ""']"")[0].title;
			if (col == 1) {
				 filterdata = themed");
            WriteLiteral(@"ata.find(x => x.LineItem == val);
				if (!!filterdata && !!filterdata[val] && filterdata[val] == ""bold"")
					cell.style.fontWeight = 'bold';
				

				
			}
			else if (col == 0) {
				
				if(!!filterdata && !!filterdata['Config_bold'] && filterdata['Config_bold'] == ""bold"")
					 cell.style.fontWeight = 'bold';
				

			}
			else {
				
				
				if (!!filterdata && !!filterdata[header] && filterdata[header] == ""bold"")
					cell.style.fontWeight = 'bold';
				if (!!filterdata && !!filterdata[header+""_color""])
						$(cell).css('color', 'blue');
					
				if (!!filterdata && !!filterdata[header+'_underline'] )
						$(cell).css('text-decoration','underline');
					
			}


				if (col == 0) {
					$(cell).css(""text-align"", ""left"");
					if (cell.innerText.indexOf("":"") >= 0) {
						cell.style.fontWeight = 'bold';
					}
				}
			}
		


	});



	function letter(nNum) {
		var a = ""A"".charCodeAt(0);
		return String.fromCharCode(a + nNum - 1);
	}

	console.log(catIn");
            WriteLiteral("dex);\r\n\tmytable.moveColumn(catIndex, 0);\r\n\tmytable.setWidth(0, 250);\r\n\tmytable.setWidth(0, 250);\r\n\r\n\r\n\tfor (var i = 1; i < longest.length; i++) {\r\n\t\tmytable.setWidth(i, 100);\r\n\t}\r\n\r\n\r\n</script>");
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
