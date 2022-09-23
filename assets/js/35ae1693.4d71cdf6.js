"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[5534],{3905:function(e,t,n){n.d(t,{Zo:function(){return u},kt:function(){return c}});var o=n(67294);function i(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function r(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){i(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,o,i=function(e,t){if(null==e)return{};var n,o,i={},a=Object.keys(e);for(o=0;o<a.length;o++)n=a[o],t.indexOf(n)>=0||(i[n]=e[n]);return i}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(o=0;o<a.length;o++)n=a[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(i[n]=e[n])}return i}var s=o.createContext({}),p=function(e){var t=o.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):r(r({},t),e)),n},u=function(e){var t=p(e.components);return o.createElement(s.Provider,{value:t},e.children)},h={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},m=o.forwardRef((function(e,t){var n=e.components,i=e.mdxType,a=e.originalType,s=e.parentName,u=l(e,["components","mdxType","originalType","parentName"]),m=p(n),c=i,d=m["".concat(s,".").concat(c)]||m[c]||h[c]||a;return n?o.createElement(d,r(r({ref:t},u),{},{components:n})):o.createElement(d,r({ref:t},u))}));function c(e,t){var n=arguments,i=t&&t.mdxType;if("string"==typeof e||i){var a=n.length,r=new Array(a);r[0]=m;var l={};for(var s in t)hasOwnProperty.call(t,s)&&(l[s]=t[s]);l.originalType=e,l.mdxType="string"==typeof e?e:i,r[1]=l;for(var p=2;p<a;p++)r[p]=n[p];return o.createElement.apply(null,r)}return o.createElement.apply(null,n)}m.displayName="MDXCreateElement"},47643:function(e,t,n){n.r(t),n.d(t,{assets:function(){return u},contentTitle:function(){return s},default:function(){return c},frontMatter:function(){return l},metadata:function(){return p},toc:function(){return h}});var o=n(87462),i=n(63366),a=(n(67294),n(3905)),r=["components"],l={title:"VBA Looping",tags:["VBA"],permalink:"/vba/looping/"},s=void 0,p={unversionedId:"vba-looping",id:"vba-looping",title:"VBA Looping",description:"The term looping refers to repeating a block of VBA statements numerous times.",source:"@site/docs/vba/19-vba-looping.md",sourceDirName:".",slug:"/vba-looping",permalink:"/vba/vba-looping",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:19,frontMatter:{title:"VBA Looping",tags:["VBA"],permalink:"/vba/looping/"},sidebar:"tutorialSidebar",previous:{title:"If-Then-Else and Select Case structure",permalink:"/vba/vba-if-then-structure-select-case"},next:{title:"Bug Finding",permalink:"/vba/vba-bug-finding"}},u={},h=[{value:"For -Next Loop",id:"for--next-loop",level:2},{value:"For-Next example",id:"for-next-example",level:3},{value:"For-Next example with an Exit For statement",id:"for-next-example-with-an-exit-for-statement",level:3},{value:"Do-While Loop",id:"do-while-loop",level:2},{value:"Do-Until Loop",id:"do-until-loop",level:2},{value:"Looping through a Collection",id:"looping-through-a-collection",level:2}],m={toc:h};function c(e){var t=e.components,n=(0,i.Z)(e,r);return(0,a.kt)("wrapper",(0,o.Z)({},m,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"The term ",(0,a.kt)("em",{parentName:"p"},"looping")," refers to repeating a block of VBA statements numerous times. "),(0,a.kt)("p",null,"VBA provides various looping command for repeating code to make correct decision making. "),(0,a.kt)("p",null,"We will go through them in following topics: "),(0,a.kt)("h2",{id:"for--next-loop"},"For -Next Loop"),(0,a.kt)("p",null,"The simplest type of loop is a ",(0,a.kt)("inlineCode",{parentName:"p"},"For-Next")," loop. Here\u2019s the syntax for this structure:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"For counter = start To end [Step stepval]\n[statements]\n[Exit For]\n[statements]\nNext [counter]\n")),(0,a.kt)("p",null,"The ",(0,a.kt)("em",{parentName:"p"},"looping")," is controlled by a counter variable, which starts at one value and stops at another value. "),(0,a.kt)("p",null,"The statements between the ",(0,a.kt)("inlineCode",{parentName:"p"},"For")," statement and the ",(0,a.kt)("inlineCode",{parentName:"p"},"Next")," statement are the statements that get repeated in the loop. "),(0,a.kt)("p",null,"To see how this works, keep reading. "),(0,a.kt)("h3",{id:"for-next-example"},"For-Next example"),(0,a.kt)("p",null,"The following example shows a ",(0,a.kt)("inlineCode",{parentName:"p"},"For-Next")," loop that doesn\u2019t use the optional Step value or the optional ",(0,a.kt)("inlineCode",{parentName:"p"},"Exit")," For statement. "),(0,a.kt)("p",null,"This routine loops 10 times and uses the VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"MsgBox")," function to show a number from 1 to 10: "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowNumbers1()\n  Dim i As Integer\n  For i = 1 to 10\n    MsgBox i\n  Next i\nEnd Sub\n")),(0,a.kt)("p",null,"In this example, ",(0,a.kt)("inlineCode",{parentName:"p"},"i")," (the loop counter variable) starts with a value of 1 and increases by 1 each time through the loop. "),(0,a.kt)("p",null,"Because I didn\u2019t specify a Step value the ",(0,a.kt)("inlineCode",{parentName:"p"},"MsgBox")," method uses the value of i as an argument. "),(0,a.kt)("p",null,"The first time through the ",(0,a.kt)("em",{parentName:"p"},"loop"),", ",(0,a.kt)("inlineCode",{parentName:"p"},"i")," is 1 and the procedure shows a number. "),(0,a.kt)("p",null,"The second time through (i = 2), the procedure show a number, and so on. "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowNumbers2()\n  Dim i As Integer Step 2\n  For i = 1 to 10\n    MsgBox i\n  Next i\nEnd Sub\n")),(0,a.kt)("p",null,"Count starts out as 1 and then takes on a value of 3, 5, 7, and 9. The final Count value is 9. "),(0,a.kt)("p",null,"The Step value determines how the counter is ",(0,a.kt)("em",{parentName:"p"},"incremented"),". Notice that the upper loop value (9) is not used because the highest value of Count after 9 would be 11, and 11 is larger than 10. "),(0,a.kt)("h3",{id:"for-next-example-with-an-exit-for-statement"},"For-Next example with an Exit For statement"),(0,a.kt)("p",null,"A ",(0,a.kt)("inlineCode",{parentName:"p"},"For-Next")," loop can also include one or more ",(0,a.kt)("inlineCode",{parentName:"p"},"Exit For")," statements within the loop. "),(0,a.kt)("p",null,"When VBA encounters this statement, the loop terminates immediately. "),(0,a.kt)("p",null,"Here\u2019s the same procedure as in the preceding section, rewritten to insert random numbers. "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowNumbers3()\n  Dim i As Integer Step 2\n  For i = 1 to 10\n    If i = 5 Then\n      MsgBox \u201cThis is a mid value\u201d\n      Exit For\n    End If\n    MsgBox i\n  Next i\nEnd Sub\n")),(0,a.kt)("p",null,"This routine performs the as earlier but when the variable i reached to 5, it shows a message, stating that this is a mid value and exit from loop. "),(0,a.kt)("h2",{id:"do-while-loop"},"Do-While Loop"),(0,a.kt)("p",null,"VBA supports another type of looping structure known as a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop. "),(0,a.kt)("p",null,"Unlike a For-Next loop, a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop continues until a specified condition is met. "),(0,a.kt)("p",null,"Here\u2019s the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop syntax:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Do-While Structure\nDo [While condition]\n  [statements]\n  [Exit Do]\n  [statements]\nLoop\n")),(0,a.kt)("p",null,"The following example uses a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop. This routine uses 1 as a starting point and runs through next numbers. "),(0,a.kt)("p",null,"The loop continues until the routine encounter the condition of ",(0,a.kt)("inlineCode",{parentName:"p"},"i = 8"),". "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Do-While Example\nSub ShowNumbers4()\n  Dim i As Integer\n  Do While i <> 8\n    MsgBox i\n    i = i + 1\n  Loop\nEnd Sub\n")),(0,a.kt)("p",null,"Some people prefer to code a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop as a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop While")," loop. "),(0,a.kt)("p",null,"This example performs exactly as the previous procedure but uses different loop syntax:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Do-Loop While Example\nSub ShowNumbers5()\n  Dim i As Integer\n  Do \n    MsgBox i\n    i = i + 1\n  Loop While i <> 8\nEnd Sub\n")),(0,a.kt)("p",null,"Here\u2019s the key difference between the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," and ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop While")," loops. "),(0,a.kt)("p",null,"The ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop always performs its conditional test first. If the test is not true, the instructions inside the loop are never executed. "),(0,a.kt)("p",null,"The ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop While")," loop, on the other hand, always performs its conditional test after the instructions inside the loop are executed. "),(0,a.kt)("p",null,"Thus, the loop instructions are always executed at least once, regardless of the test. "),(0,a.kt)("p",null,"This difference can sometimes have a big effect on how your program functions. "),(0,a.kt)("h2",{id:"do-until-loop"},"Do-Until Loop"),(0,a.kt)("p",null,"The ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop structure is similar to the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," structure. "),(0,a.kt)("p",null,"The two structures differ in their handling of the tested condition. "),(0,a.kt)("p",null,"A program continues to execute a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop while the condition remains true. "),(0,a.kt)("p",null,"In a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop, the program executes the loop until the condition is true. Here\u2019s the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," syntax: "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Do-Until Structure\nDo [Until condition]\n  [statements]\n  [Exit Do]\n  [statements]\nLoop\n")),(0,a.kt)("p",null,"The following example is the same one presented for the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop but recoded to use a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop: "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowNumbers6()\n  Dim i As Integer\n  Do Until i <> 8\n    MsgBox i\n    i = i + 1\n  Loop\nEnd Sub\n")),(0,a.kt)("p",null,"Just like with the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop, you may encounter a different form of the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop \u2014 a ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop Until")," loop. "),(0,a.kt)("p",null,"The following example, which has the same effect as the preceding procedure, demonstrates an alternate syntax for this type of loop: "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowNumbers7()\n' Do-Loop Until Example\n  Dim i As Integer\n  Do \n    MsgBox i\n    i = i + 1\n  Loop Until i <> 8\nEnd Sub\n")),(0,a.kt)("p",null,"There is a subtle difference in how the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop and the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop Until")," loop operate. "),(0,a.kt)("p",null,"In the former, the test is performed at the beginning of the loop, before anything in the body of the loop is executed. "),(0,a.kt)("p",null,"This means that it is possible that the code in the loop body will not be executed if the test condition is met. "),(0,a.kt)("p",null,"In the later version, the condition is tested at the end of the loop. "),(0,a.kt)("p",null,"Therefore, at a minimum, the ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Loop")," Until loop always results in the body of the loop being executed once. "),(0,a.kt)("p",null,"Another way to think about it is like this: The ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-While")," loop keeps looping as long as the condition is true. "),(0,a.kt)("p",null,"The ",(0,a.kt)("inlineCode",{parentName:"p"},"Do-Until")," loop keeps looping as long as the condition is False. "),(0,a.kt)("h2",{id:"looping-through-a-collection"},"Looping through a Collection"),(0,a.kt)("p",null,"VBA supports yet another type of looping \u2014 looping through each object in a ",(0,a.kt)("strong",{parentName:"p"},"collection")," of objects. "),(0,a.kt)("p",null,"Please note that I have not covered Object topic so far. For your understanding I give a brief explanation about ",(0,a.kt)("strong",{parentName:"p"},"collection"),". "),(0,a.kt)("p",null,"A ",(0,a.kt)("strong",{parentName:"p"},"collection")," is a group of same type of objects. "),(0,a.kt)("p",null,"For example, a drawing file in any CAD application is a collection of Sheets, and each sheet is a collection of drawing views and so on. "),(0,a.kt)("p",null,"When you need to loop through each object in a collection, use the For Each-Next structure. The syntax is "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' For Each-Next Structure\nFor Each element In collection\n  [statements]\n  [Exit For]\n  [statements]\nNext [element]\n")),(0,a.kt)("p",null,"The following example loops through each drawing sheet in the active drawing and shows name of each active drawing sheet: "),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' For Each-Next Example\nOption Explicit\nDim swApp As SldWorks.SldWorks\nDim swPart As SldWorks.ModelDoc2\nDim swDwg As SldWorks.DrawingDoc\nDim BoolStatus As Boolean\nDim SheetNamesList As Variant\nSub ShowSheetName()\n  Set swApp = Application.SldWorks\n  Set swPart = swApp.ActiveDoc\n  Set swDwg = swPart\n  SheetNamesList = swDwg.GetSheetNames\n  Dim SheetName As Variant\n  For Each SheetName In SheetNamesList\n    MsgBox SheetName\n  Next SheetName\nEnd Sub\n")),(0,a.kt)("p",null,"In this example, first we get the list of all sheet names in opened drawing, then we loop through each sheet name in collection and show sheet name in a message box. "),(0,a.kt)("p",null,"For this example please notes that we did not need to load all sheet, this code can work on non-activate and non-loaded sheets also. "),(0,a.kt)("p",null,"Next post will be about ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("em",{parentName:"strong"},"Bug Finding")),"."))}c.isMDXComponent=!0}}]);