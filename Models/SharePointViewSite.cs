// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SharePointViewSiteSnippet>
using Microsoft.Graph;
using System;

namespace GraphTutorial.Models
{
    public class SharePointViewSite
    {
        public string Id { get; private set; }
        public string WebUrl { get; private set; }
        public string Title { get; private set; }
        public string Description { get; private set; }

        public SharePointViewSite(Site graphSite)
        {
            Id = graphSite.Id;
            WebUrl = graphSite.WebUrl;
            Title = graphSite.DisplayName;
            Description = graphSite.Description;
        }
    }
}
// </SharePointViewSiteSnippet>
