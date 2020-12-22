(window.webpackJsonp=window.webpackJsonp||[]).push([[25],{153:function(e,t,n){"use strict";n.d(t,"a",(function(){return c})),n.d(t,"b",(function(){return m}));var o=n(0),a=n.n(o);function i(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function l(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function r(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?l(Object(n),!0).forEach((function(t){i(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):l(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function p(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},i=Object.keys(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var s=a.a.createContext({}),b=function(e){var t=a.a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):r(r({},t),e)),n},c=function(e){var t=b(e.components);return a.a.createElement(s.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return a.a.createElement(a.a.Fragment,{},t)}},h=a.a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,i=e.originalType,l=e.parentName,s=p(e,["components","mdxType","originalType","parentName"]),c=b(n),h=o,m=c["".concat(l,".").concat(h)]||c[h]||u[h]||i;return n?a.a.createElement(m,r(r({ref:t},s),{},{components:n})):a.a.createElement(m,r({ref:t},s))}));function m(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var i=n.length,l=new Array(i);l[0]=h;var r={};for(var p in t)hasOwnProperty.call(t,p)&&(r[p]=t[p]);r.originalType=e,r.mdxType="string"==typeof e?e:o,l[1]=r;for(var s=2;s<i;s++)l[s]=n[s];return a.a.createElement.apply(null,l)}return a.a.createElement.apply(null,n)}h.displayName="MDXCreateElement"},83:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return l})),n.d(t,"metadata",(function(){return r})),n.d(t,"rightToc",(function(){return p})),n.d(t,"default",(function(){return b}));var o=n(2),a=n(7),i=(n(0),n(153)),l={id:"vba-looping",title:"VBA Looping"},r={unversionedId:"vba/vba-looping",id:"vba/vba-looping",isDocsHomePage:!1,title:"VBA Looping",description:"The term looping refers to repeating a block of VBA statements numerous times.",source:"@site/docs\\vba\\2019-01-6-vba-looping.md",slug:"/vba/vba-looping",permalink:"/docs/vba/vba-looping",version:"current",sidebar:"vba",previous:{title:"If-Then-Else and Select Case structure",permalink:"/docs/vba/vba-if-else"},next:{title:"Bug Finding",permalink:"/docs/vba/vba-bug-find"}},p=[{value:"For -Next Loop",id:"for--next-loop",children:[{value:"For-Next example",id:"for-next-example",children:[]},{value:"For-Next example with an Exit For statement",id:"for-next-example-with-an-exit-for-statement",children:[]}]},{value:"Do-While Loop",id:"do-while-loop",children:[]},{value:"Do-Until Loop",id:"do-until-loop",children:[]},{value:"Looping through a Collection",id:"looping-through-a-collection",children:[]}],s={rightToc:p};function b(e){var t=e.components,n=Object(a.a)(e,["components"]);return Object(i.b)("wrapper",Object(o.a)({},s,n,{components:t,mdxType:"MDXLayout"}),Object(i.b)("p",null,"The term ",Object(i.b)("em",{parentName:"p"},"looping")," refers to repeating a block of VBA statements numerous times. "),Object(i.b)("p",null,"VBA provides various looping command for repeating code to make correct decision making. "),Object(i.b)("p",null,"We will go through them in following topics: "),Object(i.b)("h2",{id:"for--next-loop"},"For -Next Loop"),Object(i.b)("p",null,"The simplest type of loop is a ",Object(i.b)("inlineCode",{parentName:"p"},"For-Next")," loop. Here\u2019s the syntax for this structure:"),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"For counter = start To end [Step stepval]\n[statements]\n[Exit For]\n[statements]\nNext [counter]\n")),Object(i.b)("p",null,"The ",Object(i.b)("em",{parentName:"p"},"looping")," is controlled by a counter variable, which starts at one value and stops at another value. "),Object(i.b)("p",null,"The statements between the ",Object(i.b)("inlineCode",{parentName:"p"},"For")," statement and the ",Object(i.b)("inlineCode",{parentName:"p"},"Next")," statement are the statements that get repeated in the loop. "),Object(i.b)("p",null,"To see how this works, keep reading. "),Object(i.b)("hr",null),Object(i.b)("h3",{id:"for-next-example"},"For-Next example"),Object(i.b)("p",null,"The following example shows a ",Object(i.b)("inlineCode",{parentName:"p"},"For-Next")," loop that doesn\u2019t use the optional Step value or the optional ",Object(i.b)("inlineCode",{parentName:"p"},"Exit")," For statement. "),Object(i.b)("p",null,"This routine loops 10 times and uses the VBA ",Object(i.b)("inlineCode",{parentName:"p"},"MsgBox")," function to show a number from 1 to 10: "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Sub ShowNumbers1()\n  Dim i As Integer\n  For i = 1 to 10\n    MsgBox i\n  Next i\nEnd Sub\n")),Object(i.b)("p",null,"In this example, ",Object(i.b)("inlineCode",{parentName:"p"},"i")," (the loop counter variable) starts with a value of 1 and increases by 1 each time through the loop. "),Object(i.b)("p",null,"Because I didn\u2019t specify a Step value the ",Object(i.b)("inlineCode",{parentName:"p"},"MsgBox")," method uses the value of i as an argument. "),Object(i.b)("p",null,"The first time through the ",Object(i.b)("em",{parentName:"p"},"loop"),", ",Object(i.b)("inlineCode",{parentName:"p"},"i")," is 1 and the procedure shows a number. "),Object(i.b)("p",null,"The second time through (i = 2), the procedure show a number, and so on. "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Sub ShowNumbers2()\n  Dim i As Integer Step 2\n  For i = 1 to 10\n    MsgBox i\n  Next i\nEnd Sub\n")),Object(i.b)("p",null,"Count starts out as 1 and then takes on a value of 3, 5, 7, and 9. The final Count value is 9. "),Object(i.b)("p",null,"The Step value determines how the counter is ",Object(i.b)("em",{parentName:"p"},"incremented"),". Notice that the upper loop value (9) is not used because the highest value of Count after 9 would be 11, and 11 is larger than 10. "),Object(i.b)("hr",null),Object(i.b)("h3",{id:"for-next-example-with-an-exit-for-statement"},"For-Next example with an Exit For statement"),Object(i.b)("p",null,"A ",Object(i.b)("inlineCode",{parentName:"p"},"For-Next")," loop can also include one or more ",Object(i.b)("inlineCode",{parentName:"p"},"Exit For")," statements within the loop. "),Object(i.b)("p",null,"When VBA encounters this statement, the loop terminates immediately. "),Object(i.b)("p",null,"Here\u2019s the same procedure as in the preceding section, rewritten to insert random numbers. "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Sub ShowNumbers3()\n  Dim i As Integer Step 2\n  For i = 1 to 10\n    If i = 5 Then\n      MsgBox \u201cThis is a mid value\u201d\n      Exit For\n    End If\n    MsgBox i\n  Next i\nEnd Sub\n")),Object(i.b)("p",null,"This routine performs the as earlier but when the variable i reached to 5, it shows a message, stating that this is a mid value and exit from loop. "),Object(i.b)("hr",null),Object(i.b)("h2",{id:"do-while-loop"},"Do-While Loop"),Object(i.b)("p",null,"VBA supports another type of looping structure known as a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop. "),Object(i.b)("p",null,"Unlike a For-Next loop, a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop continues until a specified condition is met. "),Object(i.b)("p",null,"Here\u2019s the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop syntax:"),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' Do-While Structure\nDo [While condition]\n  [statements]\n  [Exit Do]\n  [statements]\nLoop\n")),Object(i.b)("p",null,"The following example uses a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop. This routine uses 1 as a starting point and runs through next numbers. "),Object(i.b)("p",null,"The loop continues until the routine encounter the condition of ",Object(i.b)("inlineCode",{parentName:"p"},"i = 8"),". "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' Do-While Example\nSub ShowNumbers4()\n  Dim i As Integer\n  Do While i <> 8\n    MsgBox i\n    i = i + 1\n  Loop\nEnd Sub\n")),Object(i.b)("p",null,"Some people prefer to code a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop as a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop While")," loop. "),Object(i.b)("p",null,"This example performs exactly as the previous procedure but uses different loop syntax:"),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' Do-Loop While Example\nSub ShowNumbers5()\n  Dim i As Integer\n  Do \n    MsgBox i\n    i = i + 1\n  Loop While i <> 8\nEnd Sub\n")),Object(i.b)("p",null,"Here\u2019s the key difference between the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," and ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop While")," loops. "),Object(i.b)("p",null,"The ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop always performs its conditional test first. If the test is not true, the instructions inside the loop are never executed. "),Object(i.b)("p",null,"The ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop While")," loop, on the other hand, always performs its conditional test after the instructions inside the loop are executed. "),Object(i.b)("p",null,"Thus, the loop instructions are always executed at least once, regardless of the test. "),Object(i.b)("p",null,"This difference can sometimes have a big effect on how your program functions. "),Object(i.b)("hr",null),Object(i.b)("h2",{id:"do-until-loop"},"Do-Until Loop"),Object(i.b)("p",null,"The ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop structure is similar to the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," structure. "),Object(i.b)("p",null,"The two structures differ in their handling of the tested condition. "),Object(i.b)("p",null,"A program continues to execute a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop while the condition remains true. "),Object(i.b)("p",null,"In a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop, the program executes the loop until the condition is true. Here\u2019s the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," syntax: "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' Do-Until Structure\nDo [Until condition]\n  [statements]\n  [Exit Do]\n  [statements]\nLoop\n")),Object(i.b)("p",null,"The following example is the same one presented for the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop but recoded to use a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop: "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Sub ShowNumbers6()\n  Dim i As Integer\n  Do Until i <> 8\n    MsgBox i\n    i = i + 1\n  Loop\nEnd Sub\n")),Object(i.b)("p",null,"Just like with the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop, you may encounter a different form of the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop \u2014 a ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop Until")," loop. "),Object(i.b)("p",null,"The following example, which has the same effect as the preceding procedure, demonstrates an alternate syntax for this type of loop: "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Sub ShowNumbers7()\n' Do-Loop Until Example\n  Dim i As Integer\n  Do \n    MsgBox i\n    i = i + 1\n  Loop Until i <> 8\nEnd Sub\n")),Object(i.b)("p",null,"There is a subtle difference in how the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop and the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop Until")," loop operate. "),Object(i.b)("p",null,"In the former, the test is performed at the beginning of the loop, before anything in the body of the loop is executed. "),Object(i.b)("p",null,"This means that it is possible that the code in the loop body will not be executed if the test condition is met. "),Object(i.b)("p",null,"In the later version, the condition is tested at the end of the loop. "),Object(i.b)("p",null,"Therefore, at a minimum, the ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Loop")," Until loop always results in the body of the loop being executed once. "),Object(i.b)("p",null,"Another way to think about it is like this: The ",Object(i.b)("inlineCode",{parentName:"p"},"Do-While")," loop keeps looping as long as the condition is true. "),Object(i.b)("p",null,"The ",Object(i.b)("inlineCode",{parentName:"p"},"Do-Until")," loop keeps looping as long as the condition is False. "),Object(i.b)("hr",null),Object(i.b)("h2",{id:"looping-through-a-collection"},"Looping through a Collection"),Object(i.b)("p",null,"VBA supports yet another type of looping \u2014 looping through each object in a ",Object(i.b)("strong",{parentName:"p"},"collection")," of objects. "),Object(i.b)("p",null,"Please note that I have not covered Object topic so far. For your understanding I give a brief explanation about ",Object(i.b)("strong",{parentName:"p"},"collection"),". "),Object(i.b)("p",null,"A ",Object(i.b)("strong",{parentName:"p"},"collection")," is a group of same type of objects. "),Object(i.b)("p",null,"For example, a drawing file in any CAD application is a collection of Sheets, and each sheet is a collection of drawing views and so on. "),Object(i.b)("p",null,"When you need to loop through each object in a collection, use the For Each-Next structure. The syntax is "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' For Each-Next Structure\nFor Each element In collection\n  [statements]\n  [Exit For]\n  [statements]\nNext [element]\n")),Object(i.b)("p",null,"The following example loops through each drawing sheet in the active drawing and shows name of each active drawing sheet: "),Object(i.b)("pre",null,Object(i.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"' For Each-Next Example\nOption Explicit\nDim swApp As SldWorks.SldWorks\nDim swPart As SldWorks.ModelDoc2\nDim swDwg As SldWorks.DrawingDoc\nDim BoolStatus As Boolean\nDim SheetNamesList As Variant\nSub ShowSheetName()\n  Set swApp = Application.SldWorks\n  Set swPart = swApp.ActiveDoc\n  Set swDwg = swPart\n  SheetNamesList = swDwg.GetSheetNames\n  Dim SheetName As Variant\n  For Each SheetName In SheetNamesList\n    MsgBox SheetName\n  Next SheetName\nEnd Sub\n")),Object(i.b)("p",null,"In this example, first we get the list of all sheet names in opened drawing, then we loop through each sheet name in collection and show sheet name in a message box. "),Object(i.b)("p",null,"For this example please notes that we did not need to load all sheet, this code can work on non-activate and non-loaded sheets also. "),Object(i.b)("p",null,"Next post will be about ",Object(i.b)("strong",{parentName:"p"},Object(i.b)("em",{parentName:"strong"},"Bug Finding")),"."))}b.isMDXComponent=!0}}]);