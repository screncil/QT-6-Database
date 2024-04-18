#ifndef OPENAPPS_H
#define OPENAPPS_H

#endif // OPENAPPS_H

#include <windows.h>
#include <stdio.h>

class Excel {
public:
    static bool OpenExcel() {
        // Шлях до Excel
        TCHAR excelPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE");

        STARTUPINFO si;
        PROCESS_INFORMATION pi;

        ZeroMemory(&si, sizeof(si));
        si.cb = sizeof(si);
        ZeroMemory(&pi, sizeof(pi));

        // Запускаємо Excel без вказівки файлу
        if (!CreateProcess(NULL, excelPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
            printf("CreateProcess failed (%d).\n", GetLastError());
            return false;
        }

        // Закриваємо дескриптори процесу та потоку
        // CloseHandle(pi.hProcess);
        // CloseHandle(pi.hThread);

        return true;
    }
};

class Word {
public:
    static bool OpenWord() {
        // Шлях до Word
        TCHAR wordPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE");

        STARTUPINFO si;
        PROCESS_INFORMATION pi;

        ZeroMemory(&si, sizeof(si));
        si.cb = sizeof(si);
        ZeroMemory(&pi, sizeof(pi));

        // Запускаємо Word без вказівки файлу
        if (!CreateProcess(NULL, wordPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
            printf("CreateProcess failed (%d).\n", GetLastError());
            return false;
        }

        // Закриваємо дескриптори процесу та потоку
        // CloseHandle(pi.hProcess);
        // CloseHandle(pi.hThread);

        return true;
    }
};

class Access {
public:
    static bool OpenAccess() {
        // Шлях до Access
        TCHAR accessPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\MSACCESS.EXE");

        STARTUPINFO si;
        PROCESS_INFORMATION pi;

        ZeroMemory(&si, sizeof(si));
        si.cb = sizeof(si);
        ZeroMemory(&pi, sizeof(pi));

        // Запускаємо Access без вказівки файлу
        if (!CreateProcess(NULL, accessPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
            printf("CreateProcess failed (%d).\n", GetLastError());
            return false;
        }

        // Закриваємо дескриптори процесу та потоку
        // CloseHandle(pi.hProcess);
        // CloseHandle(pi.hThread);

        return true;
    }
};
