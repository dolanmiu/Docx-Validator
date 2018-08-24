namespace DocxValidatorV1
{
    public struct ValidationError
    {
        public string Description;
        public string ErrorType;
        public string Node;
        public string Path;
        public string Part;
    }
}