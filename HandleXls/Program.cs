using System;

namespace HandleXls
{
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            //Application.Run(new Form1());
            var Form = new Form1();
            Form.Show();
        }
    }
}