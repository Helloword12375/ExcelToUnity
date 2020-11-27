using System.Collections.Generic;
using UnityEngine;
using System.IO;
using System.Xml;
using LitJson;
using Excel;
using System.Data;
using UnityEditor;
using System.Text.RegularExpressions;
using System;
using System.CodeDom;
using System.Reflection;
using System.CodeDom.Compiler;

/// <summary>
/// 创建XML表
/// </summary>
public class CreateTable : MonoBehaviour
{

    // 表头
    public static  string xmlRoot; 
    //表名
    public static  string TabeName;

    //第一行字段
    private static  string[] tableTop;

    //第二行字段
    private static string[] tableClass;

    //第三行字段
    private static string[] tableName;

    //表List-第四行之后
    private static List<string[]> tableList = new List<string[]>();

   //转化路径;
    private  static string FilePath;


    [MenuItem("Assets/BuildXml")]
    static void SetXmlFile()
    {
        string excelFile = Application.dataPath + "/ExcelToFile/ExcelFile";
        DirectoryInfo folder = new DirectoryInfo(excelFile);
        foreach (FileInfo file in folder.GetFiles("*.xlsx"))
        {

            FilePath = Application.dataPath + "/ExcelToFile/XmlFile/" + file.Name.Replace(".xlsx",".xml");
            xmlRoot= file.Name.Replace(".xlsx", "Xml_Table");
            TabeName = file.Name.Replace(".xlsx", "");
            ReadExcel(file.Name,0);
        }

    }

    [MenuItem("Assets/BuildJson")]
    static void SetJsonFile()
    {
        string excelFile = Application.dataPath + "/ExcelToFile/ExcelFile";
        DirectoryInfo folder = new DirectoryInfo(excelFile);
        foreach (FileInfo file in folder.GetFiles("*.xlsx"))
        {

            FilePath = Application.dataPath + "/ExcelToFile/JsonFile/" + file.Name.Replace(".xlsx", ".json");
            TabeName = file.Name.Replace(".xlsx", "");
            ReadExcel(file.Name,1);
        }

    }
    [MenuItem("Assets/BuildC#")]
    static void SetCsFile()
    {
        string excelFile = Application.dataPath + "/ExcelToFile/ExcelFile";
        DirectoryInfo folder = new DirectoryInfo(excelFile);
        foreach (FileInfo file in folder.GetFiles("*.xlsx"))
        {

            FilePath = Application.dataPath + "/ExcelToFile/CsFile/" + file.Name.Replace(".xlsx", ".cs");
            TabeName = file.Name.Replace(".xlsx", "");
            ReadExcel(file.Name, 2);
        }

    }

    /// <summary>
    /// 读取Excel
    /// </summary>
    static void ReadExcel(string ExcelPath,int model)
    {
        
            //excel文件位置 
            FileStream stream = File.Open(Application.dataPath + "/ExcelToFile/ExcelFile/" + ExcelPath, FileMode.Open);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet result = excelReader.AsDataSet();
            int rows = result.Tables[0].Rows.Count;
            int columns = result.Tables[0].Columns.Count;
   

        //初始化字段
        tableTop = new string[columns];
        tableClass = new string[columns];
        tableName = new string[columns];

        //存前三行行字段
        for (int i = 0; i < columns; i++)
        {
            tableTop[i] = result.Tables[0].Rows[0][i].ToString();
            tableClass[i] = result.Tables[0].Rows[1][i].ToString();
            tableName[i] = result.Tables[0].Rows[2][i].ToString();
        }

        //从第四行开始读 读信息
        for (int i = 3; i < rows; i++)
        {
            //临时表
            string[] table = new string[columns];
            //赋值表信息
            for (int j = 0; j < columns; j++)
            {
                string nvalue = result.Tables[0].Rows[i][j].ToString();
                table[j] = nvalue;
            }
            //添加到List
            tableList.Add(table);
        }
        if (model == 0)
            CreateXMLTable();
        else if (model == 1)
            CreateJSONTable();
        else if (model == 2)
            CreateCsTable();

    }


    /// <summary>
    /// 创建Xml表格
    /// </summary>
    static  void CreateXMLTable()
    {
      
        if (File.Exists(FilePath)) {
            File.Delete(FilePath);
        }

        //xml对象；
        XmlDocument xmll = new XmlDocument();
        //跟节点
        XmlElement Root = xmll.CreateElement(xmlRoot);

        for (int i = 0; i < tableList.Count; i++)
        {
            XmlElement xmlElement = xmll.CreateElement(TabeName);
            xmlElement.SetAttribute(tableTop[0], tableList[i][0]);

            for (int j = 0; j < tableTop.Length; j++)
            {
                XmlElement infoElement = xmll.CreateElement(tableTop[j]);
                infoElement.InnerText = tableList[i][j];
                xmlElement.AppendChild(infoElement);
            }
            Root.AppendChild(xmlElement);
        }

        xmll.AppendChild(Root);
        xmll.Save(FilePath);

    }
    /// <summary>
    /// 创建Json表格
    /// </summary>
    static void CreateJSONTable()
    {
        
        if (File.Exists(FilePath))
        {
            File.Delete(FilePath);
        }

        JsonData jsons = new JsonData();
        jsons.SetJsonType(JsonType.Array);
  
        for (int i = 0; i < tableList.Count; i++)
        {

            JsonData json = new JsonData();
           
            for (int j = 0; j < tableTop.Length; j++)
            {
                if(tableClass[j]=="int")
                json[tableTop[j]] =int.Parse( tableList[i][j]);
                else if (tableClass[j] == "float")
                    json[tableTop[j]] = float.Parse(tableList[i][j]);
                else if (tableClass[j] == "double")
                    json[tableTop[j]] = double.Parse(tableList[i][j]);
                else if (tableClass[j] == "string")
                    json[tableTop[j]] = tableList[i][j];
             
            }
            jsons.Add(json);

        }
        JsonData t = new JsonData();
        t[TabeName] = jsons;
        string s = t.ToJson();
         s = s.Replace("}", "}"+"\n");
        s = s.Replace(",", "," + "\n");
        Regex reg = new Regex(@"(?i)\\[uU]([0-9a-f]{4})");
        File.AppendAllText(FilePath, reg.Replace(s, delegate (Match m) { return ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString(); }));
    }
    /// <summary>
    /// 创建C#
    /// </summary>
    static void CreateCsTable() {
        if (File.Exists(FilePath))
        {
            File.Delete(FilePath);
        }
        CodeTypeDeclaration myClass = new CodeTypeDeclaration(TabeName); //生成类
        myClass.IsClass = true;
        myClass.TypeAttributes = TypeAttributes.Public;
        myClass.CustomAttributes.Add(new CodeAttributeDeclaration(new CodeTypeReference("System.Serializable")));
        for (int j = 0; j < tableClass.Length; j++)
        {
            CodeMemberField member = new CodeMemberField(GetTheType(tableClass[j]), tableTop[j]); //生成字段
            member.Attributes = MemberAttributes.Public;
            myClass.Members.Add(member); //把生成的字段加入到生成的类中
        }
        CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
        CodeGeneratorOptions options = new CodeGeneratorOptions();    //代码生成风格
        options.BracingStyle = "C";
        options.BlankLinesBetweenMembers = true;
        using (StreamWriter sw = new StreamWriter(FilePath))
        {
            provider.GenerateCodeFromType(myClass, sw, options); //生成文件
        }


    }
    private static Type GetTheType(string type)
    {
        switch (type)
        {
            case "string":
                return typeof(String);
            case "int":
                return typeof(Int32);
            case "float":
                return typeof(Single);
            case "long":
                return typeof(Int64);
            case "double":
                return typeof(Double);
            default:
                return typeof(String);
        }
    }
 }