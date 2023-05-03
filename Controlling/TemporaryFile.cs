namespace Controlling
{
    public class TemporaryFile : IDisposable
    {
        private string tempFilePath;

        public static TemporaryFile CreateCopy(string originalFilePath)
        {
            return new TemporaryFile(originalFilePath);
        }

        private TemporaryFile(string originalFilePath)
        {
            if (string.IsNullOrEmpty(originalFilePath))
            {
                throw new ArgumentNullException(nameof(originalFilePath));
            }

            if (!File.Exists(originalFilePath))
            {
                throw new FileNotFoundException($"File not found: {originalFilePath}", originalFilePath);
            }

            tempFilePath = Path.GetTempFileName();

            try
            {
                File.Copy(originalFilePath, tempFilePath, true);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error creating temporary file from {originalFilePath}.", ex);
            }
        }

        public string FilePath
        {
            get
            {
                if (string.IsNullOrEmpty(tempFilePath))
                {
                    throw new ObjectDisposedException("TemporaryFile");
                }

                return tempFilePath;
            }
        }

        public void Dispose()
        {
            if (!string.IsNullOrEmpty(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                }
                catch
                {
                    // Ignore any errors while deleting the file
                }

                tempFilePath = null;
            }
        }
    }
}
