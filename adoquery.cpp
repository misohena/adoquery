/**
 * @file
 * @brief ADO Query CLI
 * @author AKIYAMA Kouhei
 * @since 2018-12-18
 */
#import "c:\Program Files\Common Files\System\ADO\msado15.dll"  rename("EOF", "ADO_EOF")

#include <algorithm>
#include <vector>
#include <string>
#include <clocale>
#include <cstdio>

#include <ole2.h>

#include <tchar.h>


using string = std::basic_string<TCHAR>;

string escapeValue(const TCHAR *value)
{
	string result;
	result.reserve(_tcslen(value));
	while(*value){
		switch(*value){
		case _T('\\'): result += _T("\\\\"); break;
		case _T('\t'): result += _T("\\t"); break;
		case _T('\r'): result += _T("\\r"); break;
		case _T('\n'): result += _T("\\n"); break;
		default: result += *value; break;
		}
		++value;
	}
	return result;
}


/*
 * Return "%1%\t%2%\t%3%...\t%fieldCount%"
 */
string createDefaultFormat(int fieldCount)
{
	string formatStr;
	TCHAR buf[50];
	for(int i = 0; i < fieldCount; ++i){
		if(i > 0){
			formatStr.push_back(_T('\t'));
		}
		_stprintf_s(buf, _T("%%%d%%"), i + 1);
		formatStr += buf;
	}
	return formatStr;
}

bool matchChar(const TCHAR *&p, TCHAR ch)
{
	if(*p == ch){
		++p;
		return true;
	}
	else{
		return false;
	}
}

string getIdent(const TCHAR *&p)
{
	const TCHAR * const beg = p;
	while((*p >= _T('A') && *p <= _T('Z')) || (*p >= _T('a') && *p <= _T('z')) || (*p >= '0' && *p <= '9') || *p == '_'){
		++p;
	}
	return string(beg, p);
}

string getStringLiteral(const TCHAR *&p)
{
	if(!matchChar(p, _T('"'))){
		return string();
	}
	const TCHAR * const beg = p;
	while(*p != _T('"')){
		if(*p == _T('\0')){
			return string(); //error
		}
		//@todo skip \"
		++p;
	}
	const TCHAR * const end = p;
	++p;
	return string(beg, end);
}

string getStringLiteralSingleQuoted(const TCHAR *&p)
{
	if(!matchChar(p, _T('\''))){
		return string();
	}
	const TCHAR * const beg = p;
	while(*p != _T('\'')){
		if(*p == _T('\0')){
			return string(); //error
		}
		++p;
	}
	const TCHAR * const end = p;
	++p;
	return string(beg, end);
}

string getIntegerLiteral(const TCHAR *&p)
{
	const TCHAR * const beg = p;
	if(*p == '-'){
		++p;
	}
	if(!(*p >= '0' && *p <= '9')){
		return string(); //error
	}
	while(*p >= '0' && *p <= '9'){
		++p;
	}
	return string(beg, p);
}

void skipWhiteSpaces(const TCHAR *&p)
{
	while(*p == _T(' ')){
		++p;
	}
}

string formatFunctionFieldReference(const TCHAR *&p, const std::vector<_bstr_t> &fields, bool escapeOutputValue)
{
	if(!matchChar(p, _T('('))){
		return string(); //error
	}
	skipWhiteSpaces(p);
	const string funName = getIdent(p);
	if(funName.empty()){
		return string(); //error
	}
	std::vector<string> args;
	while(skipWhiteSpaces(p), !matchChar(p, _T(')'))){
		if(*p == _T('(')){
			args.push_back(formatFunctionFieldReference(p, fields, escapeOutputValue));
		}
		else if(*p == _T('"')){
			args.push_back(getStringLiteral(p));
		}
		else if(*p == _T('\'')){
			args.push_back(getStringLiteralSingleQuoted(p));
		}
		else if(*p == _T('-') || (*p >= _T('0') && *p <= _T('9'))){
			args.push_back(getIntegerLiteral(p));
		}
		else {
			return string(); //error
		}
	}

	// Execute Function
	if(funName == _T("field")){
		const std::size_t fieldIndex = std::stoi(args[0]);
		if(fieldIndex >= 0 && static_cast<std::size_t>(fieldIndex) < fields.size()){
			// %<num>%
			if(escapeOutputValue){
				return escapeValue(fields[fieldIndex]);
			}
			else {
				return string(fields[fieldIndex]);
			}
		}
		else {
			return string();
		}
	}
	else if(funName == _T("remove_prefix")){
		if(args.size() != 2){
			return string();
		}
		const string prefix = args[0];
		const string str = args[1];
		if(str.compare(0, prefix.size(), prefix) == 0){
			return args[1].substr(args[0].size());
		}
		else {
			return args[1];
		}
	}
	else{
		return string();
	}
}

/*
 * Return formatted fields string
 *
 * formatStr:
 *   %<digits>% => fields[<digits>] or escapeValue(fields[<digits>])
 *   \n => \n
 *   \r => \r
 *   \t => \t
 *   \\ => \
 *   \<char> => <char>
 */
string formatFields(const string &formatStr, const std::vector<_bstr_t> &fields, bool escapeOutputValue = false)
{
	string result;
	for(std::size_t i = 0; i < formatStr.size();){
		std::size_t first = formatStr.find_first_of(_T("%\\"), i);
		if(first == string::npos){
			first = formatStr.size();
		}
		if(first > i){
			result.append(formatStr, i, first - i);
			i = first;
		}

		if(formatStr[i] == _T('%')){
			const std::size_t last = formatStr.find_first_of(_T('%'), i + 1);
			if(last == string::npos){
				break; //syntax error
			}
			if(last == i + 1){
				// %%
				result +=_T('%');
			}
			else if(formatStr[i + 1] == _T('(')){
				const TCHAR * const beg = &formatStr[i + 1];
				const TCHAR *p = beg;
				result += formatFunctionFieldReference(p, fields, escapeOutputValue);
			}
			else{
				const int fieldIndex = std::stoi(formatStr.substr(i + 1, last - i - 1)) - 1;
				if(fieldIndex >= 0 && static_cast<std::size_t>(fieldIndex) < fields.size()){
					// %<num>%
					if(escapeOutputValue){
						result += escapeValue(fields[fieldIndex]);
					}
					else{
						result += fields[fieldIndex];
					}
				}
			}
			i = last + 1;
		}
		else if(formatStr[i] == _T('\\')){
			if(i + 1 >= formatStr.size()){
				break; //syntax error
			}
			switch(formatStr[i + 1]){
			case _T('t'): result += _T("\t"); break; //'\t'
			case _T('r'): result += _T("\r"); break; //'\r'
			case _T('n'): result += _T("\n"); break; //'\n'
			case _T('\\'): result += _T("\\"); break; //'\\'
			default: result += formatStr[i + 1]; break; //'\<char>'
			}
			i += 2;
		}
		else{
			result += formatStr[i];
			++i;
		}
	}
	return result;
}

int _tmain(int argc, TCHAR *argv[])
{
	// Parse command line
	struct Option {
		std::vector<const TCHAR *> switches;
		string *strPtr;
		bool *boolPtr;
		const TCHAR *description;
	};
	string query, conn, user, pass;
	string headerFormat, rowFormat;
	bool headerFormatSpecified = false;
	bool rowFormatSpecified = false;
	bool escapeOutputValue = false;
	bool help = false;
	Option options[] = {
		Option{{_T("conn")}, &conn, nullptr, L"Connection string (see: ADO Connection Open)"},
		Option{{_T("user")}, &user, nullptr, L"UserID (see: ADO Connection Open)"},
		Option{{_T("pass")}, &pass, nullptr, L"Password (see: ADO Connection Open)"},
		Option{{_T("query"), _T("q")}, &query, nullptr, L"Query string"},
		Option{{_T("header")}, &headerFormat, &headerFormatSpecified, L"Format for field names (default: \"%1%\\t%2%...\")"},
		Option{{_T("format")}, &rowFormat, &rowFormatSpecified, L"Format for field values (default: \"%1%\\t%2%...\")"},
		Option{{_T("escape")}, nullptr, &escapeOutputValue, L"Escape special characters(\\n \\r \\t \\) in output values"},
		Option{{_T("help"), _T("h"), _T("?")}, nullptr, &help, L"Help"},
	};

	for(int i = 1; i < argc; ++i){
		const TCHAR * const arg = argv[i];
		if(arg[0] == _T('/') || arg[0] == _T('-')){
			auto optIt = std::find_if(std::begin(options), std::end(options), [&arg](const Option &opt) {
					return std::find_if(opt.switches.begin(), opt.switches.end(), [&arg](const TCHAR *optSw) {
							return _tcsicmp(optSw, arg + 1) == 0; }) != opt.switches.end();
				});
			if(optIt == std::end(options)){
				_ftprintf(stderr, _T("Unknown option %s\n"), arg);
				return 1;
			}

			if(optIt->boolPtr){
				*(optIt->boolPtr) = true;
			}
			if(optIt->strPtr){
				if(i + 1 >= argc) {
					_ftprintf(stderr, _T("Specify value of %s option\n"), arg);
					return 1;
				}
				++i;
				optIt->strPtr->assign(argv[i]);
			}
		}
		else{
			_ftprintf(stderr, _T("Unknown option %s\n"), arg);
			return 1;
		}
	}

	if(conn.empty() || query.empty()){
		_ftprintf(stderr, _T("Specify /conn and /query options\n"));
		help = true;
	}
	if(help){
		_tprintf(_T("Options:\n"));
		for(Option &opt : options){
			for(const string &sw : opt.switches){
				if(opt.strPtr){
					_tprintf(_T("  /%s <string>\n"), sw.c_str());
				}
				else{
					_tprintf(_T("  /%s\n"), sw.c_str());
				}
			}
			_tprintf(_T("    %s\n"), opt.description);

		}
		return 1;
	}

	// Set locale
	setlocale(LC_ALL, "");

	// Initialize COM
	::CoInitialize(NULL);

	// Access DB
	bool hasError = false;
	ADODB::_ConnectionPtr connection;
	try{
		if(connection.CreateInstance(__uuidof(ADODB::Connection)) == S_OK){
			// Open
			connection->Open(conn.c_str(), user.c_str(), pass.c_str(), NULL);

			// Execute query
			ADODB::_RecordsetPtr recordset = connection->Execute(query.c_str(), NULL, ADODB::adCmdText);

			// Get field count
			const int fieldCount = recordset->Fields->GetCount();
			if(!headerFormatSpecified){ headerFormat = createDefaultFormat(fieldCount); }
			if(!rowFormatSpecified){ rowFormat = createDefaultFormat(fieldCount); }

			// Enumerate field names
			if(!headerFormat.empty()){
				std::vector<_bstr_t> fieldNames;
				fieldNames.reserve(fieldCount);
				for(int fi = 0; fi < fieldCount; ++fi){
					fieldNames.emplace_back(recordset->Fields->GetItem((long)fi)->GetName());
				}
				_tprintf(_T("%s\n"), formatFields(headerFormat, fieldNames).c_str());
			}

			// Enumerate Rows
			if(!rowFormat.empty()){
				std::vector<_bstr_t> fieldValues(fieldCount);
				while(!recordset->ADO_EOF){
					for(int fi = 0; fi < fieldCount; ++fi){
						fieldValues[fi].Assign(recordset->Fields->GetItem((long)fi)->GetValue().bstrVal);
					}
					_tprintf(_T("%s\n"), formatFields(rowFormat, fieldValues, escapeOutputValue).c_str());
					recordset->MoveNext();
				}
			}

			recordset->Close();
		}
		else{
			_ftprintf(stderr, _T("Failed to connect database\n"));
			hasError = true;
		}
	}catch(_com_error &e){
		_ftprintf(stderr, _T("Error(0x%08x):%s\n"), e.Error() , (TCHAR *)e.Description());
		hasError = true;
	}

	// close
	if(connection && connection->State != ADODB::adStateClosed){
		connection->Close();
	}

	::CoUninitialize();

	return hasError ? 1 : 0;
}
