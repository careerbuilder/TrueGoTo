/*
 Copyright 2015 CareerBuilder, LLC
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and limitations under the License.
 */

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