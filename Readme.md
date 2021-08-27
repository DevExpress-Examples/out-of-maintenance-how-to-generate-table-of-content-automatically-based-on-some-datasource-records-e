<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128610291/13.2.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E5011)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/WindowsFormsApplication1/Form1.cs) (VB: [Form1.vb](./VB/WindowsFormsApplication1/Form1.vb))
<!-- default file list end -->
# How to generate Table Of Content automatically based on some datasource records


<p>Let's consider the following scenario: we have some datasource (for example, DataTable) with two columns.</p><p>The first column contains some headers and the second column contains some text.</p><p>The requirement is to concatenate text from the second column into a single document and generate Table Of Content using values from the first column.</p><br />
<p>This example demonstrates how to achieve this functionality. The main idea of the demonstrated approach is to add text from the second column into the RichEditControl's document as separate paragraphs.</p><p>After that, you need to set the ParagraphStyle.OutlineLevel property for these paragraphs to a corresponding value and add the "TOC" document fields.</p><p>It results in creating the Table Of Content automatically.</p>

<br/>


