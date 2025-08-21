#include "CEpubData.h"

BOOL UnZip(const TCHAR* path, const TCHAR* out)
{
    BOOL bRet = FALSE;
    HRESULT hResult = S_FALSE;
    IShellDispatch* pIshellDispatch = NULL;
    Folder* pToFolder = NULL;
    BSTR bPath = SysAllocString(path);
    BSTR bOut = SysAllocString(out);
    VARIANT variantDir, variantFile, variantOpt;

    hResult = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, IID_IShellDispatch, (void**)&pIshellDispatch);
    if (SUCCEEDED(hResult))
    {
        VariantInit(&variantDir);
        variantDir.vt = VT_BSTR;
        variantDir.bstrVal = bOut;

        hResult = pIshellDispatch->NameSpace(variantDir, &pToFolder);
        if (SUCCEEDED(hResult))
        {
            Folder* pFromFolder = NULL;
            VariantInit(&variantFile);
            variantFile.vt = VT_BSTR;
            variantFile.bstrVal = bPath;
            pIshellDispatch->NameSpace(variantFile, &pFromFolder);

            FolderItems* fi = NULL;
            pFromFolder->Items(&fi);

            VariantInit(&variantOpt);
            variantOpt.vt = VT_I4;
            variantOpt.lVal = 0;

            VARIANT newV;
            VariantInit(&newV);
            newV.vt = VT_DISPATCH;
            newV.pdispVal = fi;
            hResult = pToFolder->CopyHere(newV, variantOpt);
            bRet = SUCCEEDED(hResult);

            pFromFolder->Release();
            pToFolder->Release();
        }
        pIshellDispatch->Release();
    }
    SysFreeString(bPath);
    SysFreeString(bOut);
    return bRet;
}

void Base64Encode(const char* Data, DWORDLONG size,std::string& dest)
{
    const char EncodeTable[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    unsigned char Tmp[4] = { 0 };
    int LineLength = 0;
    for (int i = 0; i < (int)(size / 3); i++)
    {
        Tmp[1] = *Data++;
        Tmp[2] = *Data++;
        Tmp[3] = *Data++;
        dest += EncodeTable[Tmp[1] >> 2];
        dest += EncodeTable[((Tmp[1] << 4) | (Tmp[2] >> 4)) & 0x3F];
        dest += EncodeTable[((Tmp[2] << 2) | (Tmp[3] >> 6)) & 0x3F];
        dest += EncodeTable[Tmp[3] & 0x3F];
        if (LineLength += 4, LineLength == 76) { dest += "\r\n"; LineLength = 0; }
    }
    //对剩余数据进行编码  
    int Mod = size % 3;
    if (Mod == 1)
    {
        Tmp[1] = *Data++;
        dest += EncodeTable[(Tmp[1] & 0xFC) >> 2];
        dest += EncodeTable[((Tmp[1] & 0x03) << 4)];
        dest += "==";
    }
    else if (Mod == 2)
    {
        Tmp[1] = *Data++;
        Tmp[2] = *Data++;
        dest += EncodeTable[(Tmp[1] & 0xFC) >> 2];
        dest += EncodeTable[((Tmp[1] & 0x03) << 4) | ((Tmp[2] & 0xF0) >> 4)];
        dest += EncodeTable[((Tmp[2] & 0x0F) << 2)];
        dest += "=";
    }
}

size_t Jpg2Base64(CStringW path, std::string& base64)
{
    std::string buffer;
    uint64_t Size = ReadFile(path, buffer);
    if (Size > 0)
    {
        base64 += "data:image/jpeg;base64,";
        Base64Encode(buffer.data(),Size , base64);
        return base64.length();
    }
    return 0;
}

size_t Png2Base64(CStringW path, std::string& base64)
{
    std::string buffer;
    uint64_t Size = ReadFile(path, buffer);
    if (Size > 0)
    {
        base64 += "data:image/png;base64,";
        Base64Encode(buffer.data(), Size, base64);
        return base64.length();
    }
    return 0;
}

uint64_t ReadFile(CStringW path, std::string& data)
{
    CAtlFile file;
    if (SUCCEEDED(file.Create(path, GENERIC_READ, FILE_SHARE_READ, OPEN_ALWAYS)))
    {
        DWORDLONG  Size = 0;
        file.GetSize(Size);
        data.resize(Size);
        if (SUCCEEDED(file.Read((void*)data.data(), Size)))
        {
            return Size;
        }
    }
    return 0;
}

uint64_t ReadFile(CStringW path, CStringW& data)
{
    std::string var;
    ReadFile(path, var);
    data = var.c_str();
    return data.GetLength();
}



CStringW GetCurrentDirectory()
{
    std::string path;
    path.resize(MAX_PATH * 4);
    GetCurrentDirectory(MAX_PATH * 2, (LPWSTR)path.data());
    return (LPWSTR)path.data();
}

CStringW GetOutDir()
{
    static CStringW out;
    if (out.GetLength() == 0)
    {
        out.SetString(GetCurrentDirectory()+_T("\\out\\"));
        CreateDirectory(out, nullptr);
    }
    return out;
}

CStringW GetTmpDir()
{
    static CStringW tmp;
    if (tmp.GetLength() == 0)
    {
        tmp.SetString(GetCurrentDirectory() + _T("\\tmp\\"));
        CreateDirectory(tmp,nullptr);
    }
    return tmp;
}

int32_t EnumerateFiles(CStringW path, std::vector<EnumerateFileTypeInfo>& info)
{
    CStringW search = path + _T("\\*.*");
    WIN32_FIND_DATAW data = { 0 };

    HANDLE hFind = FindFirstFileW(search, &data);
    if (hFind == INVALID_HANDLE_VALUE)
    {
        return -1;
    }
    do
    {
        if (data.cFileName[0] == _T('.') && (data.cFileName[1] == _T('\0') || data.cFileName[1] == _T('.')))
        {
            continue;
        }
        if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            EnumerateFiles(path + _T('\\') + data.cFileName, info);
        }
        else
        {
            CStringW ext = PathFindExtensionW(data.cFileName);
            ext.MakeLower();
            for (size_t i = 0; i < info.size(); i++)
            {
                if (info[i].type == ext)
                {
                    info[i].full_file_path.push_back(path+_T('\\') + data.cFileName);
                    break;
                }
            }
        }
    } while (FindNextFileW(hFind,&data));


    return 0;
}

CEpubData::CEpubData()
{
   
}

CEpubData::~CEpubData()
{
}

void CEpubData::SetDir(CStringW path)
{
    m_jump = false;
    m_path = path;
}

bool CEpubData::FindTocFile()
{
    std::vector<EnumerateFileTypeInfo> info;
    info.emplace_back(_T(".ncx"));
    EnumerateFiles(m_path, info);
    m_toc = info.at(0).full_file_path.at(0);
    if (m_toc.GetLength() > 0)
    {
        return true;
    }
    return false;
}

int CEpubData::GetCatalogInfo()
{
    m_catalog_title_max_length = 0;
    CStringW data;
    if (ReadFile(m_toc, data) > 0)
    {
        bool find_text = false;
        bool find_content_src = false;
        int len = data.GetLength();
        CStringW name;
        CStringW addr;
        int pos = data.Find(L"<navMap>");
        for (size_t i = pos; i < len; i++)
        {
            if ((data[i] == L'<' || data[i + 1] == L'/')
                && (data[i + 2] == L'N' || data[i + 2] == L'n')
                && (data[i + 3] == L'A' || data[i + 3] == L'a')
                && (data[i + 4] == L'V' || data[i + 4] == L'v')
                && (data[i + 5] == L'M' || data[i + 5] == L'm')
                && (data[i + 6] == L'A' || data[i + 6] == L'a')
                && (data[i + 7] == L'P' || data[i + 7] == L'p'))
            {
                break;
            }
            if (!find_text && data[i] == L'<' && data[i + 5] == L'>'
                && (data[i + 1] == L'T' || data[i + 1] == L't')
                && (data[i + 2] == L'E' || data[i + 2] == L'e')
                && (data[i + 3] == L'X' || data[i + 3] == L'x')
                && (data[i + 4] == L'T' || data[i + 4] == L't'))
            {
                name = "";
                for (size_t j = i + 6; j < len; j++)
                {
                    if (data[j] != L'<')
                    {
                        name += data[j];
                    }
                    else
                    {
                        find_text = true;
                        i = j;
                        break;
                    }
                }
            }

            if (!find_content_src && data[i] == L'<' && data[i + 8] == L' '
                && (data[i + 1] == L'C' || data[i + 1] == L'c')
                && (data[i + 2] == L'O' || data[i + 2] == L'o')
                && (data[i + 3] == L'N' || data[i + 3] == L'n')
                && (data[i + 4] == L'T' || data[i + 4] == L't')
                && (data[i + 5] == L'E' || data[i + 5] == L'e')
                && (data[i + 6] == L'N' || data[i + 6] == L'n')
                && (data[i + 7] == L'T' || data[i + 7] == L't')
                && (data[i + 9] == L'S' || data[i + 9] == L's')
                && (data[i + 10] == L'R' || data[i + 10] == L'r')
                && (data[i + 11] == L'C' || data[i + 11] == L'c'))
            {
                addr = "";
                // 9 10 11 12 13
                // s r c   =  "
                for (size_t j = i + 14; j < len; j++)
                {
                    if (data[j] != L'"')
                    {
                        addr += data[j];
                    }
                    else
                    {
                        find_content_src = true;
                        i = j;
                        break;
                    }
                }
            }
            if (find_text && find_content_src)
            {
                m_catalog.emplace_back(name, addr);
                m_catalog_title_max_length = name.GetLength() > m_catalog_title_max_length ? name.GetLength() : m_catalog_title_max_length;
                find_text = false;
                find_content_src = false;
            }
        }
    }
    return m_catalog.size();
}

int CEpubData::GetImages()
{
    std::vector<CStringW>  images;
    bool is_jpg = false;
    std::vector<EnumerateFileTypeInfo> info;
    info.emplace_back(_T(".jpg"));
    info.emplace_back(_T(".jpeg"));
    info.emplace_back(_T(".png"));
    EnumerateFiles(m_path, info);
    if (info.at(0).full_file_path.size() > 0)
    {
        images.swap(info.at(0).full_file_path);
        is_jpg = true;
    }
    if (info.at(1).full_file_path.size() > 0)
    {
        images.swap(info.at(1).full_file_path);
        is_jpg = true;
    }
    if (info.at(2).full_file_path.size() > 0)
    {
        images.swap(info.at(2).full_file_path);
        is_jpg = false;
    }
    for (size_t i = 0; i < images.size(); i++)
    {
        CPathW path = ATLPath::FindFileName(images.at(i));
        path.RemoveExtension();
        if (!ExcludeFile(path))
        {
            std::string data;
            if (is_jpg)
            {
                Jpg2Base64(images.at(i), data);
            }
            else
            {
                Png2Base64(images.at(i), data);
            }
            m_images[path] = data;
        }
    }
    return m_images.size();
}

int CEpubData::GetHtmls()
{
    std::vector<CStringW> list;
    int pos = 0;
    CStringW name;
    for (size_t i = 0; i < m_catalog.size(); i++)
    {
        pos = m_catalog.at(i).content.Find(L'#', 0);
        if (pos == -1)
        {
            if (name != m_catalog.at(i).content)
            {
                name = m_catalog.at(i).content;
                list.emplace_back(name);
                continue;
            }
        }
        else
        {
            CStringW name2(m_catalog.at(i).content,pos);
            if (name2 != name)
            {
                name = name2;
                list.emplace_back(name2);
            }
        }
    }
    
    pos = list.at(0).ReverseFind(L'/');
    CStringW mid(list.at(0).Left(pos));
    for (size_t i = 0; i < mid.GetLength(); i++)
    {
        if (mid.GetAt(i) == L'/')
        {
            mid.SetAt(i, L'\\');
        }
    }
    mid += L'\\';
    m_htmls.resize(list.size());
    CStringW root = m_toc.Left(m_toc.GetLength() - 7);
    for (size_t i = 0; i < list.size(); i++)
    {
        CStringW name = ATLPath::FindFileName(list.at(i));
        CStringW path = root +mid + name;
        ReadFile(path, m_htmls.at(i));
    }
    return m_htmls.size();
}

bool CEpubData::Start()
{
    if(SUCCEEDED(m_file.Create(OutFilePath(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS)))
    {
        WriteHead();
        Write(m_data.GetBuffer(), m_data.GetLength());
        for (size_t i = 0; i < m_htmls.size(); i++)
        {
            DisposeBody(i);
            Write(m_data.GetBuffer(), m_data.GetLength());
            Write((WCHAR*)L"++==++", wcslen(L"++==++"));
        }
        m_data += L" </body>\r\n  </html>";
        Write(m_data.GetBuffer(), m_data.GetLength());
        return true;
    }
    return false;
}

CStringW CEpubData::OutFilePath()
{
    return GetOutDir() + ATLPath::FindFileName(m_path) + _T(".html");
}

size_t CEpubData::Write(WCHAR* data, size_t len)
{
    int ulen = WideCharToMultiByte(CP_UTF8, 0, data, len, NULL, 0, NULL, NULL);
    std::wstring buffer;
    buffer.resize(ulen);
    WideCharToMultiByte(CP_UTF8, 0, data, len, (LPSTR)buffer.data(), ulen, NULL, NULL);
    m_file.Write(buffer.data(), ulen);
    m_data = "";
    return ulen;
}


bool CEpubData::ExcludeFile(CStringW name)
{
    int num = 0;
    for (size_t i = 0; i < name.GetLength(); i++)
    {
        if (name.GetAt(i) >= _T('0') && name.GetAt(i) <= _T('9'))
        {
            num++;
        }
    }
    return num < 2;
}

void CEpubData::WriteHead()
{
    m_data += L"<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">\r\n<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n";
    m_data += L"<head>";
    CStringW& data = m_htmls.at(0);
    uint64_t len = data.GetLength();
    bool find_title = false;
    bool find_meta = false;
    for (size_t i = 0; i < len; i++)
    {
        // emmmmm Title TITLE title 
        if (!find_title)
        {
            if (data[i] == L'<'
                && (data[i + 1] == L'T' || data[i + 1] == L't')
                && (data[i + 2] == L'I' || data[i + 2] == L'i')
                && (data[i + 3] == L'T' || data[i + 3] == L't')
                && (data[i + 4] == L'L' || data[i + 4] == L'l')
                && (data[i + 5] == L'E' || data[i + 5] == L'e'))
            {
                m_data += L"<title>";
                for (size_t j = i + 7; j < len; j++)
                {
                    m_data += data[j];
                    if (data[j] == L'/' && 
                       (data[j+1] == L'>' || data[j + 1] == L'T' || data[j + 1] == L't'))
                    {
                        m_data += L"title>\r\n";
                        find_title = true;
                        break;
                    }
                }
            }
        }

        // emmmmm Meta meta META
        if (!find_meta)
        {
            if (data[i] == L'<'
                && (data[i + 1] == L'M' || data[i + 1] == L'm')
                && (data[i + 2] == L'E' || data[i + 2] == L'e')
                && (data[i + 3] == L'T' || data[i + 3] == L't')
                && (data[i + 4] == L'A' || data[i + 4] == L'a'))
            {
                m_data += L"<meta ";
                for (size_t j = i + 6; j < len; j++)
                {
                    m_data += data[j];
                    if (data[j] == L'/' && data[j + 1] == L'>')
                    {
                        m_data += L">\r\n";
                        find_meta = true;
                        break;
                    }
                }
            }
        }
        if (data[i] ==  L'<' && data[i + 1] == L'/'
            && (data[i + 2] == L'h' || data[i + 1] == L'H')
            && (data[i + 3] == L'E' || data[i + 2] == L'E')
            && (data[i + 4] == L'a' || data[i + 3] == L'A')
            && (data[i + 5] == L'd' || data[i + 4] == L'D'))
        {
            break;
        }
        if (find_meta && find_title)
        {
            break;
        }
    }
    m_data += L"</head>\r\n<body>\r\n";
}

uint64_t CEpubData::GetBodyChildPos(uint32_t index)
{
    if (index > m_htmls.size())
    {
        return 0;
    }
    CStringW& data = m_htmls.at(index);
    for (size_t i = 0; i < data.GetLength(); i++)
    {
        if (data[i] == L'<'
            && (data[i + 1] == L'b' || data[i + 1] == L'B')
            && (data[i + 2] == L'o' || data[i + 2] == L'O')
            && (data[i + 3] == L'd' || data[i + 3] == L'D')
            && (data[i + 4] == L'y' || data[i + 4] == L'Y'))
        {
            for (size_t j = i+5; j < data.GetLength(); j++)
            {
                if (data[j] == L'<')
                {
                    return j;
                }
            }
        }
    }
    return 0;
}

void CEpubData::DisposeBody(uint32_t index)
{
    CStringW& data = m_htmls.at(index);
    size_t len = data.GetLength();

    WCHAR* title = nullptr;
    WCHAR* herf = nullptr;
    uint64_t last_line_break_src = 0;
    uint64_t last_line_break_dest = 0;
    bool is_not_end = false;
    bool is_write = true;
    size_t text_pos = 0;

    for (size_t i = GetBodyChildPos(index); i < len; i++)
    {
        if (data[i] == L'<')
        {
            is_not_end = true;
            text_pos = 0;
        }
        if (data[i] == L'>')
        {
            is_not_end = false;
            text_pos = i+1;
        }
        if (data[i] == L'\r\n' || (data[i] == L'\r' && data[i+1] == L'\n'))
        {
            last_line_break_src = i;
            last_line_break_dest = m_data.GetLength();
        }
        if (!is_not_end && data[i+1] == L'<')
        {
            int len = i - text_pos + 1;
            if (len > 1 && len <= m_catalog_title_max_length)
            {
                CStringW name(data.GetBuffer() + text_pos, len);
                for (size_t j = 0; j < m_catalog.size(); j++)
                {
                    if (m_catalog.at(j).title.GetLength() == name.GetLength())
                    {
                        int error_total_ = 0;
                        for (size_t k = 0; k < name.GetLength(); k++)
                        {
                            if (m_catalog.at(j).title[k] != name[k])
                            {
                                error_total_++;
                            }
                        }
                        if (error_total_ == 0)
                        {
                            if (index == 0 && j == m_catalog.size() -1)
                            {
                                std::vector<CatalogInfo>().swap(m_catalog_info);
                                m_jump = true;
                            }
                            if (m_jump)
                            {
                                m_data.Insert(last_line_break_dest, L"\r\n<hr></hr>\r\n");
                            }
                            

                            if (!m_jump)
                            {
                                m_catalog_info.emplace_back(j, last_line_break_dest);
                            }
                            break;
                        }
                    }
                }

            }
        }
        if (is_not_end && (data[i] == L's' || data[i] == L'S')
            && (data[i+1] == L'r' || data[i+1] == L'R')
            && (data[i+2] == L'c' || data[i+2] == L'C'))
        {
            CStringW img;
            // 012345
            // src="
            for (size_t j = i+5; j < len; j++)
            {
                if (data[j] != L'"')
                {
                    img += data[j];
                }
                else
                {
                    i = j;
                    break;
                }
            }
            img = ATLPath::FindFileName(img);
            ATLPath::RemoveExtension(img.GetBuffer());
            //auto item = m_images.find(img);
            is_write = false;
            m_data += L"src=\"";
            m_data += m_images[img].c_str();
            m_data += L'\"';

        }
        if (is_not_end && (data[i] == L'h' || data[i] == L'H')
            && (data[i+1] == L'r' || data[i+1] == L'R')
            && (data[i+2] == L'e' || data[i+2] == L'E')
            && (data[i+3] == L'f' || data[i+3] == L'F'))
        {
            is_write = false;
            CStringW herf;
            // 012345
            // herf="
            for (size_t j = i+6; j < len; j++)
            {
                if (data[j] != L'"')
                {
                    herf += data[j];
                }
                else
                {
                    i = j;
                    break;
                }
            }
            int aa = herf.Find(L'#', 0);
            if (aa != -1)
            {
                m_data += L"href=\"";
                m_data += herf.GetBuffer()+aa;
                m_data += L"\"";
            }
        }
        if ((data[i] == L'<' && data[i+1] == L'/')
            && (data[i + 2] == L'b' || data[i + 2] == L'B')
            && (data[i + 3] == L'o' || data[i + 3] == L'O')
            && (data[i + 4] == L'd' || data[i + 4] == L'D')
            && (data[i + 5] == L'y' || data[i + 5] == L'Y'))
        {
            break;
        }

        if (is_write)
        {
            m_data += data[i];
        }
        is_write = true;
    }
    if (!m_jump)
    {
        if (m_catalog_info.size() < m_catalog.size())
        {
            for (size_t i = 0; i < m_catalog_info.size(); i++)
            {
                m_data.Insert(m_catalog_info.at(i).pos, L"\r\n<hr></hr>\r\n");
            }
            std::vector<CatalogInfo>().swap(m_catalog_info);
            m_jump = true;
        }
    }
}