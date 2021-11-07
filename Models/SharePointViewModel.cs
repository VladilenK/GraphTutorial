// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SharePointViewModelSnippet>
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GraphTutorial.Models
{
    public class SharePointViewModel
    {
        private List<SharePointViewSite> _sites;

        public SharePointViewModel()
        {
            _sites = new List<SharePointViewSite>();
        }

        public SharePointViewModel(IEnumerable<Site> sites)
        {
            _sites = new List<SharePointViewSite>();

            if (sites != null)
            {
                foreach (var item in sites)
                {
                    _sites.Add(new SharePointViewSite(item));
                }
            }
        }

        // These properties get all sites 
        public SitesViewModel AllSites
        {
            get
            {
                return new SitesViewModel(
                  GetSitesForUser());
            }
        }


        private IEnumerable<SharePointViewSite> GetSitesForUser()
        {
            return _sites;
        }
    }
}
// </SharePointViewModelSnippet>
