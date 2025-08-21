#pragma once
#include <Windows.h>

#include <combaseapi.h>
#include <ShlDisp.h>
#include <MsHTML.h>
#import <msxml6.dll>
using namespace MSXML2;

#include <atlstr.h>
#include <atlpath.h>
#include <atlfile.h>
#include <atlcom.h>
#include <atlconv.h>
#include <atlsafe.h>


#include<iostream>
#include <initializer_list>
#include <string>
#include <vector>
#include <map>



BOOL UnZip(const TCHAR* path, const TCHAR* out);
void Base64Encode(const char* src, DWORDLONG size ,std::string&dest);
size_t Jpg2Base64(CStringW path, std::string& base64);
size_t Png2Base64(CStringW path, std::string& base64);
uint64_t ReadFile(CStringW path, std::string& data);
uint64_t ReadFile(CStringW path, CStringW& data);
CStringW GetCurrentDirectory();
CStringW GetOutDir();
CStringW GetTmpDir();


struct EnumerateFileTypeInfo
{
    CStringW type;                                          // eg: .txt                                
    std::vector<CStringW> full_file_path;
    EnumerateFileTypeInfo(CStringW type_)
    {
        type = type_;
    }
};

int32_t EnumerateFiles(CStringW path, std::vector<EnumerateFileTypeInfo>&info);


struct CatalogNode
{
    CStringW title;
    CStringW content;
    CatalogNode() {}
    CatalogNode(CStringW text_, CStringW content_)
    {
        title = text_;
        content = content_;
    }
};

struct CatalogInfo
{
    size_t id;
    size_t pos;
    CatalogInfo() {}
    CatalogInfo(size_t id_, size_t pos_)
    {
        id = id_;
        pos = pos_;
    }
};


class CEpubData
{
public:
    CEpubData();
    ~CEpubData();

    void SetDir(CStringW path);
    bool FindTocFile();
    int  GetCatalogInfo();
    int  GetImages();
    int  GetHtmls();
    bool Start();
    void Clear();
    CStringW OutFilePath();
    size_t Write(WCHAR* data, size_t len);

    // 只有1个数字 or 没有数字
    bool ExcludeFile(CStringW name);
    void WriteHead();
    uint64_t GetBodyChildPos(uint32_t index);
    void DisposeBody(uint32_t index);


private:
    CStringW                                    m_path;
    CStringW                                    m_toc;
    std::vector<CatalogNode>                    m_catalog;                          // 从toc.ncx中找到的目录名字以及对应的html文件
    int                                         m_catalog_title_max_length;
    std::map<CStringW ,std::string>             m_images;
    std::vector<CStringW>                       m_htmls;
    CAtlFile                                    m_file;
    CStringW                                    m_save_path;
    CStringW                                    m_data;                             // buffe
    std::vector<CatalogInfo>                    m_catalog_info;
    bool                                        m_jump;
    
};
