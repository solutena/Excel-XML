# Excel-XML

엑셀에서 XML을 편리하게 사용할 수 있는 기능을 제공합니다.

Git에서 엑셀이 충돌하여 병합하고 싶을 때 간단히 해결할 수 있습니다.

# VBA
```
Sub Export()
    On Error GoTo ErrorHandler
    
    Dim XmlMap As XmlMap
    Dim FileName As String
    Dim FilePath As String

    Set XmlMap = ActiveWorkbook.XmlMaps("Array")
    FileName = Split(ThisWorkbook.Name, ".")(0)
    FilePath = ThisWorkbook.Path & "\" & FileName & ".xml"

    ActiveWorkbook.SaveAsXMLData FilePath, XmlMap
    ActiveWorkbook.Save

    MsgBox FilePath & vbCrLf & "XML 데이터를 성공적으로 내보냈습니다.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "XML 데이터를 저장하는 도중 오류가 발생했습니다.", vbCritical
End Sub

Sub ImportSchema()
    Dim FileName As String
    Dim FilePath As String

    FileName = Split(ThisWorkbook.Name, ".")(0) & "_Schema"
    FilePath = ThisWorkbook.Path & "\" & FileName & ".xml"

    If Dir(FilePath) <> "" Then
        If MsgBox(FilePath & vbCrLf & vbCrLf & "스키마를 다시 불러오시겠습니까?", vbYesNo) = vbYes Then
            For i = ActiveWorkbook.XmlMaps.Count To 1 Step -1
                ActiveWorkbook.XmlMaps(i).Delete
            Next
            ActiveWorkbook.XmlMaps.Add(FilePath).Name = "Array"
        End If
    Else
        MsgBox FilePath & vbCrLf & vbCrLf & "스키마 파일이 존재하지 않습니다.", vbExclamation
    End If
End Sub

Sub ImportXML()
    Dim XmlMap As XmlMap
    Dim FileName As String
    Dim FilePath As String
    
    FileName = Split(ThisWorkbook.Name, ".")(0)
    FilePath = ThisWorkbook.Path & "\" & FileName & ".xml"
    
    If MsgBox(FilePath & vbCrLf & vbCrLf & "XML을 불러와 데이터에 덮어쓰시겠습니까?", vbYesNo) = vbYes Then
        Set XmlMap = ActiveWorkbook.XmlMaps("Array")
        XmlMap.Import Url:=FilePath
    End If
End Sub
```

## 함수 설명

`Export`

"엑셀파일이름" 이름으로 된 XML을 내보낸 후

엑셀을 저장 합니다.

`ImportSchema`

"엑셀파일이름_Schema" 이름으로 된 스키마를 읽어옵니다.

개발도구 > 원본 에서 새롭게 매핑을 해주시면 됩니다. 


`ImportXML`

"엑셀파일이름" 이름으로 된 XML을 읽어와 데이터에 덮어씁니다.

git에서 엑셀이 컴플릿이 났을 때

올바르게 병합된 XML으로부터 데이터를 덮어 쓸 수 있습니다.

## 사용법

![image](https://github.com/solutena/Excel-XML/assets/22467083/b0024164-6254-44df-b432-64c07d258ef0)

Visual Basic에 들어가 모듈에 VBA 함수를 추가 합니다.

![image](https://github.com/solutena/Excel-XML/assets/22467083/747d20f6-2ef6-4577-a5f7-296f9ac51c35)

데이터를 작성 후 버튼을 만들어

각 버튼의 매크로 지정에 함수를 적용하면 됩니다.

XML로 추출할 데이터의 예제입니다.

# Unity

## XML 예제
```
using System.Xml.Serialization;
 
public class Message
{
    [XmlAttribute] public string key = string.Empty;
    [XmlAttribute] public string text = string.Empty;
}
```

XML로 추출할 데이터의 예제입니다.

## Unity-XML
```
const string xmlExtension = ".xml";
public static readonly string streamingPath = Application.streamingAssetsPath + "/";

public static void ExportXML<T>() where T : new()
{
	T[] target = { new T(), new T() };
	string key = typeof(T).FullName + "_Schema";
	string path = streamingPath + key + xmlExtension;
	var serializer = new XmlSerializer(typeof(T[]));
	using XmlTextWriter writer = new(path, Encoding.UTF8);
	writer.Formatting = Formatting.Indented;
	serializer.Serialize(writer, target);
	Debug.Log("내보내기(" + key + ") : " + path);
}

public static T[] ImportXML<T>()
{
	string key = typeof(T).FullName;
	string path = streamingPath + key + xmlExtension;
	Debug.Log("불러오기(" + key + ") : " + path);
	var serializer = new XmlSerializer(typeof(T[]));
	using StringReader stringReader = new(File.ReadAllText(path));
	return (T[])serializer.Deserialize(stringReader);
}
```

`ExportXML`

"클래스명_Schema" 이름으로 된 XML스키마를 내보냅니다.

스키마는 2개의 데이터로된 배열로 만들어 내보냅니다.

배열로 만들어야 엑셀에서 표의 형태로 생성됩니다.

`ImportXML`

"클래스명" 이름으로 된 XML을 데이터 배열로 불러옵니다.
