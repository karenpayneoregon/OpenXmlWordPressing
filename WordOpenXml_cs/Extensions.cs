using System;

namespace WordOpenXml_cs
{
    /// <summary>
    /// Code by Karen Payne MVP
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Convert pixel to EMU
        /// </summary>
        /// <param name="sender"></param>
        /// <returns></returns>
        /// <remarks>Used for properly sizing an image in a document.</remarks>
        public static int PixelToEmu(this int sender) => (int)Math.Round((decimal)sender * 9525);
    }
}
