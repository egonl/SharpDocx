using DocumentFormat.OpenXml;

namespace ActiveWord
{
    public class CodeBlock : CodePart
    {
        public string Code { get; set; }

        public int Id { get; set; }

        public void AddElement(OpenXmlElement element)
        {
            if (!this.Elements.Contains(element))
            {
                this.Elements.Add(element);
            }
        }

        //public void Execute(BaseDocument doc)
        //{
        //    var preText = this.StartText.Text.Substring(0, this.StartTagIndex);
        //    var postText = this.StartText.Text.Substring(this.StartTagIndex);
        //    //this.StartText.Text = $"{preText}*{postText}";
        //    this.StartText.Text = preText;
        //    doc.ExecuteScript(this.Id, this.StartText);
        //    this.StartText.Text += postText;
        //}
    }
}