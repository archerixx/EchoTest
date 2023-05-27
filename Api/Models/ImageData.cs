using NPOI.XWPF.UserModel;

namespace Api.Models
{
    public class ImageData
    {
        public string Name { get; set; }
        public Stream Stream { get; set; }
        public PictureType Type { get; set; }
        public int Height { get; set; }
        public int Width { get; set; }
    }
}
