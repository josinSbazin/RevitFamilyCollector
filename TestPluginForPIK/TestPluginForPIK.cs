using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Xml.Linq;
using System.Diagnostics;
using System.IO;

[TransactionAttribute(TransactionMode.Manual)]
public class TestPluginForPIK : IExternalCommand
{

    private Document doc;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        doc = commandData.Application.ActiveUIDocument.Document;
        //Для пользовательских семейств
        FilteredElementCollector userFamilies = GetFilteredEllementCollectorOfUserFamilies();

        Dictionary<string, List<Family>> dictOfUserFamiliesByCategories = GetDictOfUserFamiliesByCategories(userFamilies);
        XElement userFamiliesXml = GetUserFamiliesXelement(dictOfUserFamiliesByCategories);

        //Для системных семейств
        FilteredElementCollector systemFamilies = GetFilteredEllementCollectorOfSystemFamilies();

        Dictionary<string, List<ElementType>> dictOfSystemFamiliesByCategories = GetDictOfSystemFamiliesByCategories(systemFamilies);
        XElement systemFamiliesXml = GetSystemFamiliesXelement(dictOfSystemFamiliesByCategories);

        MergeAndSaveXml(userFamiliesXml, systemFamiliesXml, "D:/ResultForPIKbyRHanza.xml");
        
        return Result.Succeeded;
    }

    //Для пользовательских семейств
    private FilteredElementCollector GetFilteredEllementCollectorOfUserFamilies()
    {
        if (doc == null) throw new FileNotFoundException();
        return new FilteredElementCollector(doc).OfClass(typeof(Family));
    }

    //Для пользовательских семейств
    private Dictionary<string, List<Family>> GetDictOfUserFamiliesByCategories(FilteredElementCollector families)
    {
        Dictionary<string, List<Family>> result = new Dictionary<string, List<Family>>();

        foreach (Family fam in families)
        {
            string key = fam.FamilyCategory.Name;
            List<Family> listOfFamilies;

            if (!result.TryGetValue(key, out listOfFamilies))
            {
                listOfFamilies = new List<Family>();
                result.Add(key, listOfFamilies);
            }

            listOfFamilies.Add(fam);
        }

        return result;
    }

    //Для пользовательских семейств
    private XElement GetUserFamiliesXelement(Dictionary<string, List<Family>> dict)
    {
        XElement xmlRootFamilyElement = new XElement("Пользовательские_семейства");

        foreach (KeyValuePair<string, List<Family>> pair in dict)
        {
            XElement category = new XElement("Категории", pair.Key);

            foreach (Family f in pair.Value)
            {
                XElement familyName = new XElement("Семейства", f.Name);
                category.Add(familyName);

                foreach (ElementId id in f.GetFamilySymbolIds())
                {
                    XElement familyTypeName = new XElement("Типоразмеры", doc.GetElement(id).Name);
                    familyName.Add(familyTypeName);
                }
            }

            xmlRootFamilyElement.Add(category);
        }

        return xmlRootFamilyElement;
    }

    //Для системных семейств
    private FilteredElementCollector GetFilteredEllementCollectorOfSystemFamilies()
    {
        if (doc == null) throw new FileNotFoundException();
        FilteredElementCollector systemFamilies = new FilteredElementCollector(doc).WhereElementIsElementType();
        ElementClassFilter filterInv = new ElementClassFilter(typeof(FamilySymbol), true);
        systemFamilies.WherePasses(filterInv);
        return systemFamilies;
    }

    //Для системных семейств
    private Dictionary<string, List<ElementType>> GetDictOfSystemFamiliesByCategories(FilteredElementCollector families)
    {
        Dictionary<string, List<ElementType>> result = new Dictionary<string, List<ElementType>>();

        foreach (Element e in families)
        {
            ElementType type = e as ElementType;

            if (null != type)
            {
                string key = type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();
                List<ElementType> listOfFamilies;

                if (!result.TryGetValue(key, out listOfFamilies))
                {
                    listOfFamilies = new List<ElementType>();
                    result.Add(key, listOfFamilies);
                }

                listOfFamilies.Add(type);
            }
        }

        return result;
    }

    //Для системных семейств
    private XElement GetSystemFamiliesXelement(Dictionary<string, List<ElementType>> dict)
    {
        XElement xmlRootFamilyElement = new XElement("Системные_семейства");

        foreach (KeyValuePair<string, List<ElementType>> pair in dict)
        {
            XElement category = new XElement("Категории-семейства", pair.Key);

            foreach (ElementType eType in pair.Value)
            {
                XElement familyName = new XElement("Типоразмеры", eType.Name);
                category.Add(familyName);
            }

            xmlRootFamilyElement.Add(category);
        }

        return xmlRootFamilyElement;
    }


    private void MergeAndSaveXml(XElement one, XElement two, string filePath)
    {
        XElement XmlRoot = new XElement("Перечень_семейств_в_проекте");
        XmlRoot.Add(one);
        XmlRoot.Add(two);

        XDocument xmldoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"),
        new XComment("Тестовое_задание_для_ПИК_от_Ганзы_Р.А_"), XmlRoot);

        xmldoc.Save(filePath);
    }

}
