using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GISClient;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.EditorExt;
using ESRI.ArcGIS.CartoUI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using Microsoft.Win32;
using System.Drawing;
using System.Drawing.Drawing2D;
using ESRI.ArcGIS.Geometry;

namespace MapElementChecker
{
  class Program
  {
    [STAThread()]
    static void Main(string[] args)
    {
      if (args.Count() != 1) // there can be only one
      {
        WriteUsage();
        return;
      }

      string mxdPath;
      var possiblePath = args[0].Trim();
      possiblePath = possiblePath.Trim('"');
      if (possiblePath.EndsWith(@"\"))
        possiblePath = possiblePath.Substring(0, possiblePath.Length - 1);
      if (Directory.Exists(possiblePath))
      {
        mxdPath = possiblePath;
      }
      else
      {
        WriteUsage();
        return;
      }
      
      ESRI.ArcGIS.RuntimeManager.BindLicense(ESRI.ArcGIS.ProductCode.Desktop, ESRI.ArcGIS.LicenseLevel.GeodatabaseUpdate);

      var mxdFiles = Directory.GetFiles(mxdPath, "*.mxd", SearchOption.AllDirectories).OrderBy(x => x).ToArray();

      foreach (var mxdFile in mxdFiles)
      {
        IMapDocument mapDoc = new MapDocumentClass();
        mapDoc.Open(mxdFile);
        FindElements(mapDoc, mxdFile);
      }
    }
    static void FindElements(IMapDocument mapDoc, string mapPath)
    {
      Console.WriteLine("Processing map document path: " + mapPath);
      int mapCount = mapDoc.MapCount;
      IGraphicsContainer graphicsContainer = null;
      bool needsSaved = false;

      for (int i = 0; i < mapCount; i++)
      {
        IMap map = mapDoc.Map[i];
        Console.WriteLine("Processing data frame: " + map.Name);

        List<string> elementClasses = new List<string>();
        List<IElement> toBeDeleted = new List<IElement>();

        graphicsContainer = map.BasicGraphicsLayer as IGraphicsContainer;
        ProcessGraphicsContainer(graphicsContainer, elementClasses, toBeDeleted);

        IActiveView activeView = map as IActiveView;
        activeView.Activate(0);

        graphicsContainer.Reset();

        foreach (IElement element in toBeDeleted)
        {
          needsSaved = true;
          graphicsContainer.DeleteElement(element);
        }

        // debugging helper
        // List<string> elementClassesUnique = elementClasses.Distinct().ToList();
        // elementClassesUnique.ForEach(Console.WriteLine);
      }

      if (needsSaved)
      {
        var newPath = mapPath.Replace(".mxd", "_updated.mxd");
        mapDoc.SaveAs(newPath);
        Console.WriteLine("Incompatible elements found. The content of map document {0} was updated and saved as {1}.", mapPath, newPath);
      }
      else
      {
        Console.WriteLine("Map document {0} was processed and no changes were needed.", mapPath);
      }

    }
    static void ProcessGraphicsContainer(IGraphicsContainer graphicsContainer, List<string> elementClasses, List<IElement> toBeDeleted)
    {
      if (graphicsContainer == null)
      {
        Console.WriteLine("Null graphics container");
        return;
      }

      graphicsContainer.Reset();
      IElement element = graphicsContainer.Next();
      while (element != null)
      {
        IGroupElement groupElement = element as IGroupElement;
        if (groupElement != null)
        {
          ProcessGroupElement(groupElement, elementClasses, toBeDeleted);
        }
        else
        {
          ProcessElement(element, elementClasses, toBeDeleted);
        }

        element = graphicsContainer.Next();
      }
    }

    static void ProcessGroupElement(IGroupElement groupElement, List<string> elementClasses, List<IElement> toBeDeleted)
    {
      if (groupElement == null)
        return;

      int elementCount = groupElement.ElementCount;
      for (int i = 0; i < elementCount; i++)
      {
        IElement element = groupElement.Element[i];
        IGroupElement groupElementInternal = element as IGroupElement;
        if (groupElementInternal != null)
        {
          ProcessGroupElement(groupElementInternal, elementClasses, toBeDeleted);
        }
        else
        {
          ProcessElement(element, elementClasses, toBeDeleted);
        }
      }
    }
    private static void ProcessElement(IElement element, List<string> elementClasses, List<IElement> toBeDeleted)
    {
      IDisplacementLinkElement displacementLinkElement = element as IDisplacementLinkElement;
      IIdentityLinkElement identityLinkElement = element as IIdentityLinkElement;
      IDataGraphTElement dataGraphTElement = element as IDataGraphTElement;
      if (displacementLinkElement != null || identityLinkElement != null || dataGraphTElement != null)
      {
        toBeDeleted.Add(element);
      }
      IPersistStream persistStream = element as IPersistStream;
      Guid thisClassID = new Guid();
      persistStream.GetClassID(out thisClassID);
      elementClasses.Add(thisClassID.ToString());
    }

    static void WriteUsage()
    {
      Console.WriteLine("Usage: MapElementChecker.exe <inputdirectorywithmxds>");
      Console.WriteLine("Example: MapElementChecker.exe c:\\temp\\mymxds\\");
    }
  }
}
