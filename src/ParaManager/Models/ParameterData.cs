using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.ComponentModel;

namespace ParaManager.Models
{
    public class ParameterData
    {
        public string Name { get; set; }
        public string ParameterType { get; set; }  // Always string for display/serialization
        public ForgeTypeId ParameterTypeId { get; set; }  // Store ForgeTypeId for API calls
#if REVIT2025
        public ForgeTypeId Group { get; set; }
#else
        public BuiltInParameterGroup Group { get; set; }
#endif
        public bool IsInstance { get; set; }
        public bool IsShared { get; set; }
        public string GUID { get; set; }
        public List<string> Categories { get; set; }
        public string Description { get; set; }
        public bool UserModifiable { get; set; }
        public bool Visible { get; set; }
        public bool HideWhenNoValue { get; set; }

        public string GroupName
        {
            get
            {
#if REVIT2025
                if (Group == null || string.IsNullOrEmpty(Group.TypeId)) return "Invalid";
                try { return LabelUtils.GetLabelForGroup(Group); } catch { return Group.TypeId; }
#else
                return Group.ToString();
#endif
            }
        }

        public ParameterData()
        {
            Categories = new List<string>();
            UserModifiable = true;
            Visible = true;
            HideWhenNoValue = false;
        }
    }

    public class SharedParameterData
    {
        public string Name { get; set; }
        public string Group { get; set; }
        public string ParameterType { get; set; }
        public string GUID { get; set; }
        public string Description { get; set; }
        public bool UserModifiable { get; set; }
        public bool Visible { get; set; }
    }

    public class FamilyParameterData : INotifyPropertyChanged
    {
        private bool _isSelected = false;

        public string Name { get; set; }
        public string ParameterType { get; set; }
        public bool IsInstance { get; set; }
        public bool IsShared { get; set; }
        public string GUID { get; set; }
        public string Formula { get; set; }
#if REVIT2025
        public ForgeTypeId Group { get; set; }
#else
        public BuiltInParameterGroup Group { get; set; }
#endif

        /// <summary>
        /// For ListView checkbox binding
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }
        }

        /// <summary>
        /// Display: "Instance" or "Type"
        /// </summary>
        public string ScopeDisplay => IsInstance ? "Instance" : "Type";

        /// <summary>
        /// Display: "Shared" or "Family"
        /// </summary>
        public string KindDisplay => IsShared ? "Shared" : "Family";

        public string GroupName
        {
            get
            {
#if REVIT2025
                if (Group == null || string.IsNullOrEmpty(Group.TypeId)) return "Invalid";
                try { return LabelUtils.GetLabelForGroup(Group); } catch { return Group.TypeId; }
#else
                return Group.ToString();
#endif
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    /// <summary>
    /// Preset for saving/loading parameter selections
    /// </summary>
    public class ParameterSelectionPreset
    {
        public string Name { get; set; }
        public List<string> SelectedParameterNames { get; set; } = new List<string>();
        public string CreatedDate { get; set; }
    }
}
