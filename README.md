# SharpDocx
[![NuGet](https://img.shields.io/nuget/v/SharpDocx.svg)](https://www.nuget.org/packages/SharpDocx/)
[![MIT](https://img.shields.io/github/license/mashape/apistatus.svg)](https://github.com/egonl/SharpDocx/blob/master/LICENSE)
[![AppVeyor](https://img.shields.io/appveyor/ci/egonl/SharpDocx.svg)](https://ci.appveyor.com/project/egonl/SharpDocx/branch/master)

*C# based template engine for generating Word documents*

Generating documents with SharpDocx is a two step process. First you create a view in Word. A view is a Word document which also contains C# code. Code can be inserted anywhere, e.g. <%= DateTime.Now %> would insert the current date and time.

The next step is to create documents based on this view. This requires two lines of code:
```
var document = DocumentFactory.Create("view.cs.docx");
document.Generate("output.docx");
```

Out of the box SharpDocx supports inserting text, tables, images and more. See the Tutorial sample.

If you want, you can specify a view model to be used in your view. Then you could write things like <% foreach (var item in Model.MyList) { %>. See the Model sample.

If you want to do something that's not supported by SharpDocx, you can do so by creating your own document subclass. See the Inheritance example.
