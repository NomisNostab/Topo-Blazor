﻿using Microsoft.AspNetCore.Components;
using Topo.Model.Login;
using Topo.Services;

namespace Topo.Controller
{
    public class LoginController : ComponentBase
    {
        [Inject]
        public ILoginService _loginService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public LoginPageViewModel model = new LoginPageViewModel();

        internal async Task LogInClick ()
        {
            var authenticationResult = await _loginService.LoginAsync(model.Branch, model.MemberNumber, model.Password);
            model.LoginErrorMessage = "";
            if (authenticationResult != null && authenticationResult.AuthenticationSuccessResultModel.AuthenticationResult != null)
            {
                await _loginService.GetUserAsync();
                await _loginService.GetProfilesAsync();
                var isLeader = false;
                if (_storageService.GetProfilesResult != null && _storageService.GetProfilesResult.profiles != null && _storageService.GetProfilesResult.profiles.Length > 0)
                {
                    _storageService.MemberName = _storageService.GetProfilesResult.profiles[0].member?.name ?? "";
                    _storageService.GroupName = _storageService.GetProfilesResult.profiles[0].group?.name ?? "";
                    var profiles = _storageService.GetProfilesResult.profiles.ToList();
                    foreach (var profile in profiles)
                    {
                        if (profile.group != null && profile.group.roles != null && profile.group.roles.Any(x => x == "group-leader" || x.StartsWith("support-leader-group-")))
                        {
                            isLeader = true;
                            break;
                        }
                        if (profile.unit != null && profile.unit.roles != null && profile.unit.roles.Any(x => x == "unit-council"))
                        {
                            isLeader = true;
                            break;
                        }
                    }
                    _storageService.IsYouthMember = !isLeader;
                }
                _storageService.Groups = _loginService.GetGroups();
                _storageService.GroupId = _storageService.Groups.Count == 1 ? _storageService.Groups.FirstOrDefault().Key : "";
                _storageService.GroupName = _storageService.Groups.Count == 1 ? _storageService.Groups.FirstOrDefault().Value : "";

                NavigationManager.NavigateTo("index");
            }
            if (authenticationResult != null && authenticationResult.AuthenticationErrorResultModel != null)
            {
                model.LoginErrorMessage = authenticationResult.AuthenticationErrorResultModel.message;
            }
        }
    }
}
