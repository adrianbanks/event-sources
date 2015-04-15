using System;

namespace N1self.C1selfVisualStudioExtension.ActiveWindow
{
    public class ForegroundWindowInfo : IEquatable<ForegroundWindowInfo>
    {
        public string Path { get; set; }
        public string WindowTitle { get; set; }

        public bool Equals(ForegroundWindowInfo other)
        {
            if (other == null)
            {
                return false;
            }

            return string.Equals(Path, other.Path)
                   && string.Equals(WindowTitle, other.WindowTitle);
        }
    }
}
