// Guids.cs
// MUST match guids.h
using System;

namespace Careerbuilder.TrueGoTo
{
    static class GuidList
    {
        public const string guidTrueGoToPkgString = "6070a2ca-c143-4ca8-bfb1-577b52c91d6b";
        public const string guidTrueGoToCmdSetString = "5e126e63-5786-4d94-9710-7789a3114f69";

        public static readonly Guid guidTrueGoToCmdSet = new Guid(guidTrueGoToCmdSetString);
    };
}