How to alternate colours when changes day
google apps script, javascript, programming, cloud

<p>I am continuing the development of my script to auto-transfer Calendar entries into a spreadsheet. In this article I show you a convenient function to change the colour of the row when changes the day of the event, useful when for example you have more than one shifts and you want to improve the readibility of the table.</p>

<p>In comparison to the my original script that even now I use as it is more complete, I am looking for simplify the output, in order to avoid a lot of columns that do some calculations. In this case I use a formula to convert the date in an integer:</p>
<pre>
=(DATE(YEAR(i3);MONTH(i3);DAY(i3))-DATE(YEAR(i3);1;0))
</pre>

<p>Integer that I use to make a comparison whith the previous one, so I change the colour if it is not equal or mantain the same if equal. This is the code:</p>
<pre>
// Variables used to alternate colours
var columnColorCalc = 28;
var color = firstColor;
var FirstWorkingDay = sheet.getRange(firstRowDate,columnColorCalc).setFormula('=(DATE(YEAR(A' +firstRowDate+ ');MONTH(A' +firstRowDate+ ');DAY(A' +firstRowDate+ '))-DATE(YEAR(A' +firstRowDate+ ');1;0))').getValue();

// Code to improve the stilish of the table
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,1).setNumberFormat("-DD/MM-").setHorizontalAlignment("center");
    sheet.getRange(i,3,totalRows,2).setNumberFormat("HH:mm");

    // Code to alternate colours
    var workingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue();
    if( FirstWorkingDay == workingDay ){
        sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    } else if (color == firstColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = secondColor; sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    } else if (color == secondColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = firstColor; sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    }
     // Code to alternate colours

}
</pre>

<p>Once obtained the formatting, I have instructed the script to delete the column with another for cicle:</p>
<pre>
// Clear columns added for calculations
for (var i=firstRowDate; i <= totalRows; i+=1){
    // The column used to change the colors
    sheet.getRange(i,columnColorCalc).clear();
}
</pre>

<p>I realise that is a really forcing solution, but I haven't yet known how to save a variable without using two funtions: getRange and setFormula. If someone more expert than me can help me, I will sure appreciate!</p>

<p>This is the final result:</p>

<h3>Download</h3>

<p>You can find and test my Script on GitHub at the following link:</p>

<ul>
	<li><span class="speciale">GitHub: <a href="https://github.com/rainnic/Javascript_scripts/tree/master/export_gcal_to_gsheet" target="_blank" title="My script on GitHub">../Javascript_scripts/export_gcal_to_gsheet/</a></span></li>
</ul>

<p>If you want to install it, you can follow my previous article <a href="../en/node/408" target="_blank">5 steps to auto-transfer your Calendar entries into a spreadsheet</a> or see the video wiht the entire process step by step:</p>

<p class="text-align-center"><iframe allow="autoplay; encrypted-media" allowfullscreen="" frameborder="0" height="315" src="https://www.youtube.com/embed/eEcFB99jCGg?rel=0" width="560"></iframe></p>

<p><strong>P.S.</strong> This is the first video with my voice, so I want to apologize for my bad grammar, pronunciation and some embarrassment.</p>
