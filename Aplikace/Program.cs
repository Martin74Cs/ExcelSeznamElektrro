// See https://aka.ms/new-console-template for more information
using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Microsoft.Office.Interop.Excel;
using System;
using Exc = Microsoft.Office.Interop.Excel;

var Ele = new ElektroLoad();

//Kopírování informací do revize 7
//Ele.Elektro();

//Vytvoření nového dokumentu
Ele.NovyExcel();