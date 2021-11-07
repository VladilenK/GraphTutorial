// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SitesViewModelSnippet>
using System;
using System.Collections.Generic;

namespace GraphTutorial.Models
{
    public class SitesViewModel
    {
        public IEnumerable<SharePointViewSite> Sites { get; private set; }

        public SitesViewModel(IEnumerable<SharePointViewSite> sites)
        {
            Sites = sites;
        }
    }
}
// </SitesViewModelSnippet>
