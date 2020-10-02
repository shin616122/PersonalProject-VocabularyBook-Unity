using System.Collections;
using UnityEngine;
using SFB;

public class FileBrowser : MonoBehaviour
{
    public CanvasManager canvasManager;
    private string _path;

    public void OnFileBrowserClick()
    {
        var extensions = new[] {
                new ExtensionFilter("Excel Files", "xlsx" ),
                new ExtensionFilter("All Files", "*" ),
            };
        WriteResult(StandaloneFileBrowser.OpenFilePanel("Open File", "", extensions, true));
        canvasManager.LoadFileFromPath();
    }

    public void WriteResult(string[] paths)
    {
        if (paths.Length == 0)
        {
            return;
        }

        canvasManager.FilePath = paths[0];
    }
}
