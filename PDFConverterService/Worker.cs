using System.Diagnostics;
using System.Xml;
using System.Configuration;

namespace PDFConverterService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {

                string applicationFolder = System.Configuration.ConfigurationManager.AppSettings["applicationFolder"];
                string outputFolder = System.Configuration.ConfigurationManager.AppSettings["outputFolder"];
                string processFolder = System.Configuration.ConfigurationManager.AppSettings["processFolder"];
                string successFolder = System.Configuration.ConfigurationManager.AppSettings["successFolder"];
                string hotFolder = System.Configuration.ConfigurationManager.AppSettings["hotFolder"];

                string[] files = Directory.GetFiles(hotFolder);


                CheckAndCreateDirectory(outputFolder);
                CheckAndCreateDirectory(processFolder);
                CheckAndCreateDirectory(successFolder);
                CheckAndCreateDirectory(hotFolder);


                if (files.Length > 0)
                {
                    string firstFileName = Path.GetFileName(files[0]);
                    Console.WriteLine("First file name: " + firstFileName);
                
                    string targetFilePath = Path.Combine(processFolder, Path.GetFileName(firstFileName));

                    MoveFile(hotFolder + @"\" + firstFileName, targetFilePath);

                    string CmdLine = "python " + applicationFolder + " \"" + targetFilePath + "\" \"" + outputFolder + @"\"+ firstFileName + "\"";

                    ConvertPDFtoWord(CmdLine);


                    //File goes to success
                    
                    files = Directory.GetFiles(processFolder);
                    firstFileName = findFirstFileName(files, firstFileName);



                    MoveFile(processFolder + @"\" + firstFileName, successFolder + @"\" + firstFileName);

                }
                else
                {
                    Console.WriteLine("No files found in the folder.");
                }


                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                await Task.Delay(1000, stoppingToken);
            }
        }


        public void MoveFile(string targetPath, string destinationPath)
        {
            try
            {
                // Move the file to the target location
                File.Move(targetPath, destinationPath);

                Console.WriteLine("File moved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Folder is empty: " + ex.Message);
            }
        }

        public string findFirstFileName(string[] files, string firstFileName)
        {
            if (files.Length > 0)
            {
                firstFileName = Path.GetFileName(files[0]);
                return firstFileName;
            }
            else
            {
                return null;
            }
        }

        public void ConvertPDFtoWord(string cmdLine)
        {
            Process process = new Process();

            ProcessStartInfo startInfo = new ProcessStartInfo("cmd.exe")
            {
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            process.StartInfo = startInfo;
            process.Start();

            process.StandardInput.WriteLine(cmdLine);

            process.StandardInput.Close();

            string output = process.StandardOutput.ReadToEnd();
        }

        static void CheckAndCreateDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
                Console.WriteLine("Directory created: " + directory);
            }
        }
    }
}