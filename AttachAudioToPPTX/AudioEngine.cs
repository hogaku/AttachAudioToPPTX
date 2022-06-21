using System.IO;
using AngleSharp.Common;
using DocumentFormat.OpenXml.Office2010.Drawing;
using ShapeCrawler;
namespace Engine;
public class AudioEngine
{
    private string homeDir;
    private string audioPath;
    private string targetpptxFileName;
    public AudioEngine(string audioPath, string targetFileName)
    {
        var homeDir = Directory.GetCurrentDirectory();
        string[] homeDirAry = homeDir.Split("\\");
        List<string> homePaths = homeDirAry.ToList();
        homePaths.RemoveAt(9);
        homePaths.RemoveAt(8);
        homePaths.RemoveAt(7);
        homeDir = String.Join("\\",homePaths.ToArray());
        this.audioPath = homeDir + audioPath;
        this.targetpptxFileName = targetFileName;
        try
        {
            //Set the current directory.
            Directory.SetCurrentDirectory(homeDir);

        }
        catch (DirectoryNotFoundException e)
        {
            Console.WriteLine("The specified directory does not exist. {0}", e);
        }

    }
    public void AddAudioShape()
    {

        DirectoryInfo di = new DirectoryInfo(audioPath);
        FileInfo[] audioFileLists = di.GetFiles("*", SearchOption.AllDirectories); // ディレクトリ以下のすべてのサブディレクトリのファイルの一覧を取得する

        foreach (FileInfo f in audioFileLists)
        {
            Console.WriteLine($"{f.FullName}");
        }
        // Add audio
        try
        {
            var targetFilePathtargetFilePath = homeDir + targetpptxFileName;
            Console.WriteLine("読み込み先ファイル名:{0}", targetpptxFileName);

            IPresentation presentation = SCPresentation.Open(targetpptxFileName, isEditable: true);
            int scount = presentation.Slides.Count;
            Console.WriteLine("読み込み済みpptxのスライド枚数:{0}", scount.ToString());

            foreach (var slide in presentation.Slides.Select((value, index) => new { value, index }))
            {
                Console.WriteLine("Targeting Slide Number: {0}", slide.index + 1);

                IShapeCollection shapes = presentation.Slides[slide.index].Shapes;
                using Stream mp3Stream = File.OpenRead(audioFileLists[slide.index].FullName);
                bool isAttachedAudio = presentation.Slides[slide.index].Shapes.Where(s => s.ShapeType == 0).Any(); // empty:false, any:true
                if (!isAttachedAudio) // 音声ファイルが付与されていない場合
                {
                    IAudioShape addedAudioShape = shapes.AddNewAudio(xPixel: 0, yPixels: presentation.SlideHeight, mp3Stream);
                }
            }
            presentation.Save();
            presentation.Close();
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }


    }

}