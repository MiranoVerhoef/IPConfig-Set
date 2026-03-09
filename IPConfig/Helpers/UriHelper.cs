﻿using System;
using System.Diagnostics;

namespace IPConfig.Helpers;

public static class UriHelper
{
    public static string NormalizeUri(string uri)
    {
        return new Uri(uri, UriKind.RelativeOrAbsolute).AbsoluteUri;
    }

    public static void OpenUri(string uri)
    {
        uri = NormalizeUri(uri);

        var psi = new ProcessStartInfo {
            FileName = uri,
            UseShellExecute = true
        };

        Process.Start(psi);
    }
}
