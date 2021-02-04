Public Class Help

    Function PrintHelpText()

        Console.WriteLine("Usage: DigiStripComicMatcher.exe SourcePath DestinationPath [options]" + vbNewLine +
                           +vbNewLine +
                           "All files in the Source Path, and subdirectories, will be scanned and based on their CRC32 hash be moved to the Destination Path." + vbNewLine +
                           "Available options:" + vbNewLine +
                           "-Help                Display this text" + vbNewLine +
                           "-RemoveDuplicate     Remove a file from the source directory if it already exists in the destination, based on its CRC32 hash.")

    End Function

End Class
