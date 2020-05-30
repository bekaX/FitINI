# FitINI
A simple and convenient library for INI files

简体中文版: [README_chs.md](README_chs.md "ReadMe")

[TOC]

## About the INI file
The `.ini` file is an easy-to-use way to store data. All data is recorded as key-value pairs, and annotations(comments) are supported.

Generally an `.ini` file contains several Sections(also called blocks), each with its own key-value pair, which can be accessed by the Key Name to the unique Value.

**FitINI** encapsulates common operations into a library that makes it easy and fast to read, modify and write `.ini` files. And there is some support for non-standard INI style files.

## Examples(Visual Basic.Net)
Load a standard `.ini` file, modify it and save it to another file

    ' Load file
    Dim ini = FitINI.INI.LoadFromFile("config.ini")
    ' Read content
    Dim sec As FitINI.Section = ini.Section("SecName1")
    Debug.WriteLine(sec.Item("KeyName1"))
    Debug.WriteLine(ini.Entry("SecName2", "KeyName2"))
    ' Modify
    sec = ini.Add("SecName3")
    sec.Item("KeyName3") = "Hello"
    ' Save to another file
    ini.SaveToFile("config2.ini")

## Some simple API usage instructions(Visual Basic.Net)

### Classes
**INI**: Indicates an INI object

**Section**: Represents a section inside an INI object

**LoadOption**: Used to specify options when parsing INI file data

The name of the section is included in the INI object and not in the Section object.
All classes are not thread safe. If sharing in multiple threads, it is recommended to use the *SyncLock* keyword.

---

Thanks for using :) ~

Currently maintained by **beka** from MFS
