// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using GraphTutorial.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TimeZoneConverter;

namespace GraphTutorial.Controllers
{
    public class SharePointController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<HomeController> _logger;

        public SharePointController(
            GraphServiceClient graphClient,
            ILogger<HomeController> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
        }

        // <IndexSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Sites.Read.All" })]
        public async Task<IActionResult> Index()
        {
            try
            {
                var sites = await GetUserFollowSites();
                var model = new SharePointViewModel(sites);

                return View(model);
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw;
                }

                return View(new SharePointViewModel())
                    .WithError("Error getting SharePoint view", ex.Message);
            }
        }
        // </IndexSnippet>

        // <SharePointNewGetSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Sites.ReadWrite" })]
        public IActionResult New()
        {
            return View();
        }
        // </SharePointNewGetSnippet>

        // <SharePointNewPostSnippet>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [AuthorizeForScopes(Scopes = new[] { "Sites.ReadWrite" })]

        // <GetSharePointViewSnippet>
        private async Task<IList<Site>> GetUserFollowSites()
        {
            // Configure a SharePoint view for the current week

            var sites = await _graphClient.Me.FollowedSites
                .Request()
                // Get max 50 per request
                .Top(50)
                // Only return fields app will use
                .Select(e => new
                {
                    e.Id,
                    e.DisplayName,
                    e.WebUrl,
                    e.Description
                })
                // Order results chronologically
                // .OrderBy("webUrl")
                .GetAsync();

            IList<Site> allSites;
            // Handle case where there are more than 50
            if (sites.NextPageRequest != null)
            {
                allSites = new List<Site>();
                // Create a page iterator to iterate over subsequent pages
                // of results. Build a list from the results
                var pageIterator = PageIterator<Site>.CreatePageIterator(
                    _graphClient, sites,
                    (e) =>
                    {
                        allSites.Add(e);
                        return true;
                    }
                );
                await pageIterator.IterateAsync();
            }
            else
            {
                // If only one page, just use the result
                allSites = sites.CurrentPage;
            }

            return allSites;
        }

        // </GetSharePointViewSnippet>
    }
}
