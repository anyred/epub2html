#include "CEpubData.h"

int main()
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (SUCCEEDED(hr))
	{
		SHFILEOPSTRUCTA FileOp;
		FileOp.fFlags = FOF_NOCONFIRMATION;
		FileOp.hNameMappings = NULL;
		FileOp.hwnd = NULL;
		FileOp.lpszProgressTitle = NULL;
		FileOp.pFrom = ".\\tmp";
		FileOp.pTo = NULL;
		FileOp.wFunc = FO_DELETE;
		SHFileOperationA(&FileOp);
		FileOp.pFrom = ".\\out";
		SHFileOperationA(&FileOp);

		std::vector<CStringW> list;
		std::vector<EnumerateFileTypeInfo> info;
		info.emplace_back(_T(".epub"));
		EnumerateFiles(GetCurrentDirectory(), info);
		size_t size = info[0].full_file_path.size();
		for (size_t i = 0; i < size; i++)
		{
			CPathW Name = ATLPath::FindFileName(info[0].full_file_path[i].GetString());
			Name.RemoveExtension();
			CStringW uncompress = GetTmpDir() + Name.m_strPath;
			CreateDirectoryW(uncompress, nullptr);
			list.emplace_back(uncompress);

			CPathW NewName = info[0].full_file_path[i].GetString();
			NewName.RenameExtension(_T(".zip"));
			MoveFileW(info[0].full_file_path[i].GetString(),NewName);

			UnZip(NewName, uncompress);
			Name.RenameExtension(_T(".epub"));
			MoveFileW(NewName, info[0].full_file_path[i].GetString());
		}

		for (size_t i = 0; i < list.size(); i++)
		{
			CEpubData data;
			data.SetDir(list.at(i));
			data.FindTocFile();
			data.GetCatalogInfo();
			data.GetImages();
			data.GetHtmls();
			data.Start();
		}
		CoUninitialize();
	}
	return 0;
}