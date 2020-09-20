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

/**
* Converts file with serz.exe
* 
* @param[in] fpath the pointer to file path to convert.
* @param[in] serzPath the pointer to file path of serz app.
*/
void convertWithSerz(fs::path* fpath, fs::path* serzPath) {
    //builds and run string 'serz.exe "<filename to convert>" > nul'
    system((serzPath->string() + " \"" + fpath->string() + "\" > nul").c_str());

    return;
}

/**
* Formats node content with <newline> and <space>.
* 
* @param[in,out] text the pointer to wstring with node content to be formatted.
* @param[in,out] spaceCounter the pointer to counter with points from last space inserted.
* @param[in,out] newLineCounter the pointer to counter with points from last nl inserted.
*/
void formatNodeContent(wstring* text, int* spaceCounter, int* newLineCounter) {
    //increment both counters
    (*spaceCounter)++;
    (*newLineCounter)++;

    if (*newLineCounter >= 8) { //if more than 8 points from last nl, insert nl and reset counter
        *newLineCounter = 0;
        *spaceCounter = 0;
        *text += L"\n";
    }
    else if (*spaceCounter >= 2) {  //if more than 2 points from last space, insert space and reset counter 
        *spaceCounter = 0;
        *text += L" ";
    }
}

/**
* Decodes text formed terrain point into float.
* 
* @param[in] encodedTerrainPoint the pointer to wstring with text encoded terrain point.
* 
* @returns float encoded terrainPoint.
*/
float decodeTerrainPoint(wstring* encodedTerrainPoint) {
    //swap bytes order in terrain point
    for (unsigned n = 0; n < 4; n += 2) {
        swap(*(encodedTerrainPoint+n), *(encodedTerrainPoint+(6-n)));
        swap(*(encodedTerrainPoint+n+1), *(encodedTerrainPoint+(7-n)));
    }
    _bstr_t terrainPoint(encodedTerrainPoint->c_str());

    //declare union to hold both float and int
    union { float fval; std::uint32_t ival; };

    //perform some non legal memory magick
    ival = strtoul((char*)terrainPoint, 0, 16);

    return fval;
}

/**
* Encodes float into text form.
* 
* @param[in] decodedTerrainPoint the float encoded terrainPoint.
* 
* @returns text form of inputted float.
*/
wstring encodeTerrainPoint(float decodedTerrainPoint) {
    //declare union for memory magicks
    union { float fval; std::uint32_t ival; };
    fval = decodedTerrainPoint;

    //define and use ostringstream to convert union to hex
    std::ostringstream stm;
    stm << std::hex << std::uppercase << ival;

    //convert ostringstream to string
    string s_shiftedPoint = stm.str();

    //swap bytes order
    for (unsigned n = 0; n < 4; n += 2) {
        swap(s_shiftedPoint[n], s_shiftedPoint[6 - n]);
        swap(s_shiftedPoint[n + 1], s_shiftedPoint[7 - n]);
    }

    //build and return wstring
    return wstring(s_shiftedPoint.begin(), s_shiftedPoint.end());
}

/**
* Shifts terrain point by inputted value.
* 
* @param[in,out] terrainPoints the pointer to BSTRing with terrainPoints.
* @param[in] zShift the pointer to float with value to move points by.
*/
void shiftTerrainPoints(BSTR* terrainPoints, float* zShift) {
    //check if terrainPoints is allocated, should always be
    assert(terrainPoints != nullptr);
    wstring s_terrainPoints(*terrainPoints);

    //sanitaze terrainPoints for parsing purposes
    s_terrainPoints.erase(remove(s_terrainPoints.begin(), s_terrainPoints.end(), ' '), s_terrainPoints.end());
    s_terrainPoints.erase(remove(s_terrainPoints.begin(), s_terrainPoints.end(), '\n'), s_terrainPoints.end());

    wstring s_outTerrainPoints;

    //define formatting counters
    int spaceCounter = 0;
    int newLineCounter = 0;

    //for each terrain point (8 chars) decode, shift, encode, format
    for (unsigned i = 0; i < s_terrainPoints.length(); i += 8) {
        wstring s_terrainPoint = s_terrainPoints.substr(i, 8);

        s_outTerrainPoints += encodeTerrainPoint(decodeTerrainPoint(&s_terrainPoint) + *zShift);

        formatNodeContent(&s_outTerrainPoints, &spaceCounter, &newLineCounter);
    }

    //allocate string with s_outTerrainPoints data
    *terrainPoints = SysAllocString(s_outTerrainPoints.data());
}

/**
* Performs shifting on terrain file.
* 
* @param[in] terFilePath the pointer to terrainFile path.
* @param[in] zShift the pointer to float with z shift ammount.
*/
void shiftTerrainFile(fs::path* terFilePath, float* zShift) {
    //load XML
    HRESULT hr = CoInitialize(NULL);
    CComPtr<IXMLDOMDocument> iXMLDoc;
    HRESULT rootRes = iXMLDoc.CoCreateInstance(__uuidof(DOMDocument));
    VARIANT_BOOL bSuccess = false;
    iXMLDoc->load(CComVariant(terFilePath->string().c_str()), &bSuccess);

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

        //Get node name
        BSTR nodeName;
        iHgtFTileNode->get_nodeName(&nodeName);

        //If node name == d:blob get its content, perform shifting on it and write it back
        if (0 == wcscmp(nodeName, L"d:blob")) {
            readContent = ((HRESULT)iHgtFTileNode->get_text(&nodeContent) == 0);
            shiftTerrainPoints(&nodeContent, zShift);
            iHgtFTileNode->put_text(nodeContent);
            break;
        }
    }

    //Save shifted terrain file back to xml
    BSTR targetFilePath = SysAllocString(terFilePath->c_str());
    VARIANT targetFPathVariant;
    targetFPathVariant.vt = VT_BSTR;
    targetFPathVariant.bstrVal = targetFilePath;
    HRESULT res = iXMLDoc->save(targetFPathVariant);

    //Free mem
    SysFreeString(nodeContent);
    SysFreeString(targetFilePath);
}

/**
* Iterates terrain folder.
* 
* @param[in] terrainDir the pointer to path of folder with terrain tiles.
* @param[in] serzPath the pointer to path of serz app.
* @param[in] zShift the pointer to float with z shift ammount.
*/
void iterateTerrainFiles(fs::path* terrainDir, fs::path* serzPath, float* zShift) {
    //get backup directory
    fs::path backupDir(terrainDir->parent_path() / "Backup" / "Terrain");
    error_code code;
    //create or empty existing backup directory
    fs::create_directory(backupDir.parent_path()); ;
    if (!fs::create_directory(backupDir, code)) {
        fs::remove_all(backupDir);
        fs::create_directory(backupDir);
    }

    //convert each terrainTile file to xml, perform shifting and convert back to bin file
    for (const auto& entry : fs::directory_iterator(*terrainDir)) {
        fs::path filePath = entry.path();
        cout << "Now shifting \"" << filePath.string() << "\"..." << endl;
        if (!entry.is_directory() && filePath.extension() == ".bin") {
            convertWithSerz(&filePath, serzPath); //convert bin to xml
            fs::rename(filePath.string(), backupDir / filePath.filename()); //backup old bin file
            filePath.replace_extension(".xml");
            shiftTerrainFile(&filePath, zShift); //shift xml file
            convertWithSerz(&filePath, serzPath); //convert back to bin
            fs::remove(filePath); //remove temporary xml
        }
        cout << "Shifting \"" << filePath.string() << "\" done!" << endl;
    }
}

/**
* Iterates route folder and checks for Terrain folder.
* 
* @param[in] routeDir the pointer to path of route root folder.
* @param[in] serzPath the pointer to path of serz app.
* @param[in] zShift the pointer to float with z shift ammount.
*/
void iterateRouteFolder(fs::path* routeDir, fs::path* serzPath, float* zShift) {
    for (const auto& entry : fs::directory_iterator(*routeDir)) {
        fs::path ePath = entry.path();
        if (entry.is_directory() && ePath.string().find("Backup") == string::npos) { //if is directory and not in backup directory
            string wdir = ePath.filename().string();
            if (wdir == "Terrain") { //if is Terrain directory, perform shifting
                cout << "Starting terrain shifting..." << endl;
                iterateTerrainFiles(&ePath, serzPath, zShift);
                cout << "Terrain shifting done!" << endl;
            }
        }
    }
}

/**
 * Checks serz presence at such path and exits program if is not present.
 *
 * @param[in] serzPath the path with serz.exe.
 */
void checkSerz(fs::path serzPath) {
    if (fs::exists(serzPath))
        return;
    cout << "Serz.exe was not found at \"" << serzPath.string() << "\"!" << endl << "Check route path validity, or check Serz.exe entered path!" << endl;
    exit(-1);
}

int main()
{
    TCHAR NPath[MAX_PATH]; //declares NPath with maximal length of path
    GetCurrentDirectory(MAX_PATH, NPath); //gets current wdir
    fs::path routeDir(NPath);
    fs::path serzPath = routeDir.parent_path().parent_path().parent_path() / "serz.exe"; //creates serz exe path from current wdir
    checkSerz(serzPath); //check if serz exists at such path
    //gets desired z offset from user
    float zOffset;
    cout << "Please input z offset:" << endl;
    cin >> zOffset;
    cout << "Shifting all terrain tiles by " << zOffset << "!" << endl;
    iterateRouteFolder(&routeDir, &serzPath, &zOffset); //iterates route folder for terrainTile files and performs shifting
    system("pause"); //waits for user input
}
