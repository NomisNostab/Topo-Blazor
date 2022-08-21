using Microsoft.AspNetCore.Components;
using Topo.Model.Members;
using Topo.Services;

namespace Topo.Controller
{

    public class MembersController : ComponentBase
    {
        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        public MembersPageViewModel model = new MembersPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            model.Units = _storageService.Units;
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId; 
            _storageService.UnitId = model.UnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Value == model.UnitId).FirstOrDefault().Value;
            var allMembers = await _membersService.GetMembersAsync(model.UnitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            model.UnitName = _storageService.UnitName;
        }
    }
}
