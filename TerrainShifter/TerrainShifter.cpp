#include <iostream>
#include <filesystem>
#include <assert.h>
#include <msxml.h>
#include <atlbase.h>
#include <string>
#include <comdef.h>
#include <sstream>

using namespace std;
namespace fs = filesystem;

void convertWithSerz(fs::path fpath, fs::path* serzPath) {
    system((serzPath->string() + " \"" + fpath.string() + "\" > nul").c_str());

    return;
}

void formatNodeContent(wstring* text, int* spaceCounter, int* newLineCounter) {
    (*spaceCounter)++;
    (*newLineCounter)++;
    if (*newLineCounter >= 8) {
        *newLineCounter = 0;
        *spaceCounter = 0;
        *text += L"\n";
    }
    else if (*spaceCounter >= 2) {
        *spaceCounter = 0;
        *text += L" ";
    }
}

void shiftTerrainPoints(BSTR* terrainPoints, float zShift) {
    assert(terrainPoints != nullptr);
    wstring s_terrainPoints(*terrainPoints);

    s_terrainPoints.erase(remove(s_terrainPoints.begin(), s_terrainPoints.end(), ' '), s_terrainPoints.end());
    s_terrainPoints.erase(remove(s_terrainPoints.begin(), s_terrainPoints.end(), '\n'), s_terrainPoints.end());

    wstring outTerrainPoints;

    int spaceCounter = 0;
    int newLineCounter = 0;

    for (unsigned i = 0; i < s_terrainPoints.length(); i += 8) {
        wstring _terrainPoint = s_terrainPoints.substr(i, 8);

        for (unsigned n = 0; n < 4; n+=2) {
            swap(_terrainPoint[n], _terrainPoint[6 - n]);
            swap(_terrainPoint[n + 1], _terrainPoint[7 - n]);
        }
        _bstr_t terrainPoint (_terrainPoint.c_str());

        union { float fval; std::uint32_t ival; };
        ival = strtoul((char*)terrainPoint, 0, 16);
        fval += zShift;

        std::ostringstream stm;
        stm << std::hex << std::uppercase << ival;
        string s_shiftedPoint = stm.str();

        for (unsigned n = 0; n < 4; n += 2) {
            swap(s_shiftedPoint[n], s_shiftedPoint[6 - n]);
            swap(s_shiftedPoint[n + 1], s_shiftedPoint[7 - n]);
        }
        outTerrainPoints += wstring(s_shiftedPoint.begin(), s_shiftedPoint.end());
        formatNodeContent(&outTerrainPoints, &spaceCounter, &newLineCounter);
    }

    *terrainPoints = SysAllocString(outTerrainPoints.data());
}

void shiftTerrainFile(fs::path terFilePath, float zShift) {
    //load XML
    HRESULT hr = CoInitialize(NULL);
    CComPtr<IXMLDOMDocument> iXMLDoc;
    HRESULT rootRes = iXMLDoc.CoCreateInstance(__uuidof(DOMDocument));
    VARIANT_BOOL bSuccess = false;
    iXMLDoc->load(CComVariant(terFilePath.string().c_str()), &bSuccess);

    // Get a pointer to the root
    CComPtr<IXMLDOMElement> iRootElm;
    iXMLDoc->get_documentElement(&iRootElm);

    //Get pointer to Record node
    CComPtr<IXMLDOMNode> iRecElement;
    iRootElm->get_firstChild(&iRecElement);

    //Get pointer to cHeightFieldTile node
    CComPtr<IXMLDOMNode> iHgtFTile;
    iRecElement->get_firstChild(&iHgtFTile);

    //Get pointer to cHeightFieldTile child nodes
    CComPtr<IXMLDOMNodeList> iHgtFTileNodes;
    iHgtFTile->get_childNodes(&iHgtFTileNodes);

    long length;
    iHgtFTileNodes->get_length(&length);

    BSTR nodeContent = BSTR();
    bool readContent = false;

    for (int i = 0; i < length; i++) {
        CComPtr<IXMLDOMNode> iHgtFTileNode;
        //Get pointer to each cHeightFieldTile child node
        iHgtFTileNodes->get_item(i, &iHgtFTileNode);

        BSTR nodeName;
        iHgtFTileNode->get_nodeName(&nodeName);

        if (0 == wcscmp(nodeName, L"d:blob")) {
            readContent = ((HRESULT)iHgtFTileNode->get_text(&nodeContent) == 0);
            shiftTerrainPoints(&nodeContent, zShift);
            iHgtFTileNode->put_text(nodeContent);
            break;
        }
    }

    BSTR targetFilePath = SysAllocString(terFilePath.c_str());
    VARIANT targetFPathVariant;
    targetFPathVariant.vt = VT_BSTR;
    targetFPathVariant.bstrVal = targetFilePath;
    HRESULT res = iXMLDoc->save(targetFPathVariant);

    SysFreeString(nodeContent);
    SysFreeString(targetFilePath);
}

void iterateTerrainFiles(fs::path routeDir, fs::path* serzPath, float zShift) {
    fs::path backupDir(routeDir.parent_path() / "Backup" / "Terrain");
    error_code code;
    fs::create_directory(backupDir.parent_path()); ;
    if (!fs::create_directory(backupDir, code)) {
        fs::remove_all(backupDir);
        fs::create_directory(backupDir);
    }
    for (const auto& entry : fs::directory_iterator(routeDir)) {
        fs::path filePath = entry.path();
        cout << "Now shifting \"" << filePath.string() << "\"..." << endl;
        if (!entry.is_directory() && filePath.extension() == ".bin") {
            convertWithSerz(filePath, serzPath);
            fs::rename(filePath.string(), backupDir / filePath.filename());
            filePath.replace_extension(".xml");
            shiftTerrainFile(filePath, zShift);
            convertWithSerz(filePath, serzPath);
            fs::remove(filePath);
        }
        cout << "Shifting \"" << filePath.string() << "\" done!" << endl;
    }

    return;
}

void iterateRouteFolder(fs::path* routeDir, fs::path* serzPath, float zShift) {
    for (const auto& entry : fs::directory_iterator(*routeDir)) {
        if (entry.is_directory() && entry.path().string().find("Backup") == string::npos) {
            string wdir = entry.path().filename().string();
            if (wdir == "Terrain") {
                cout << "Starting terrain shifting..." << endl;
                iterateTerrainFiles(entry.path(), serzPath, zShift);
                cout << "Terrain shifting done!" << endl;
            }
        }
        else if (entry.path().extension() == ".bin") {
            convertWithSerz(entry.path(), serzPath);
        }
    }

    return;
}

void checkSerz(fs::path serzPath) {
    if (fs::exists(serzPath))
        return;
    cout << "Serz.exe was not found at \"" << serzPath.string() << "\"!" << endl << "Check route path validity, or check Serz.exe entered path!" << endl;
    exit(-1);
}

int main()
{
    TCHAR NPath[MAX_PATH];
    GetCurrentDirectory(MAX_PATH, NPath);
    fs::path routeDir(NPath);
    fs::path serzPath = routeDir.parent_path().parent_path().parent_path() / "serz.exe";
    checkSerz(serzPath);
    float zOffset;
    cout << "Please input z offset:" << endl;
    cin >> zOffset;
    cout << "Shifting all terrain tiles by " << zOffset << "!" << endl;
    iterateRouteFolder(&routeDir, &serzPath, zOffset);
    system("pause");
}
