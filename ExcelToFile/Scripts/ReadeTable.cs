using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using LitJson;
using System.IO;
using System.Xml;

public class ReadeTable : MonoBehaviour
{
    public class PlayerManagers {
           public List<PlayerManager> PlayerManager = new List<PlayerManager>();
    }
    
    private  string path;
    public bool isReaderjson;
    public bool isReaderxml;
    // Start is called before the first frame update
    void Start()
    {
        if (isReaderjson)
        {
            path = Application.dataPath + "/ExcelToFile/JsonFile/PlayerManager.json";
            StreamReader streamreader = new StreamReader(path);//读取数据，转换成数据流
            JsonReader js = new JsonReader(streamreader);//再转换成json数据
            PlayerManagers r = JsonMapper.ToObject<PlayerManagers>(js);
            for (int i = 0; i < r.PlayerManager.Count; i++)
            {
                Debug.Log("json-"+r.PlayerManager[i].Id + r.PlayerManager[i].Name + r.PlayerManager[i].LeveID);
            }
        }
        if (isReaderxml)
        {
            path = Application.dataPath + "/ExcelToFile/XmlFile/PlayerManager.xml";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            XmlNodeList node = xmlDoc.SelectSingleNode("PlayerManagerXml_Table").ChildNodes;
            //遍历节点
            foreach (XmlElement x1 in node)
            {
                Debug.Log("xml-" + x1.InnerText);
                //if (x1.GetAttribute("Id") == "1")
                // {
                //foreach (XmlElement data1 in x1.ChildNodes)
                //    {
                //        Debug.Log("xml-"+data1.InnerText);
                       
                //    }
                //}
            }
        }
    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
