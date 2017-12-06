# SharpDocx
*C# library for creating Word documents*

SharpDocx is inspired by Web technologies like ASP.NET. Developers familiar with ASP.NET MVC or Web Forms should feel right at home.

First you create a view in Word. A view is a Word document which also contains C# code. The next step is to create documents based on this view. If you want, you can supply the view with a model. This requires two lines of code:
```
var document = DocumentFactory.Create("view.cs.docx", model);
document.Generate("output.docx");
```

Out of the box SharpDocx supports inserting text, tables and images. See the Tutorial sample. If you require something more specific, you can do so by creating your own document subclass (see the Inheritance example).

