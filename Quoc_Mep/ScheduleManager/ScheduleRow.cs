using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ScheduleManager
{
    public class ScheduleRow : INotifyPropertyChanged
    {
        private string _scheduleName;
        private string _elementId;
        private string _fieldName;
        private string _value;
        private int _rowIndex;
        private int _columnIndex;
        private bool _isModified;

        public string ScheduleName
        {
            get => _scheduleName;
            set { _scheduleName = value; OnPropertyChanged(); }
        }

        public string ElementId
        {
            get => _elementId;
            set { _elementId = value; OnPropertyChanged(); }
        }

        public string FieldName
        {
            get => _fieldName;
            set { _fieldName = value; OnPropertyChanged(); }
        }

        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(); }
        }

        public int RowIndex
        {
            get => _rowIndex;
            set { _rowIndex = value; OnPropertyChanged(); }
        }

        public int ColumnIndex
        {
            get => _columnIndex;
            set { _columnIndex = value; OnPropertyChanged(); }
        }

        public bool IsModified
        {
            get => _isModified;
            set { _isModified = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
            }
        }

        public void AddValue(string fieldName, string value)
        {
            _values[fieldName] = value;
            _originalValues[fieldName] = value;
        }
        
        public Element GetElement() => _element;

        public bool IsModified => _originalValues.Any(kvp => _values.ContainsKey(kvp.Key) && _values[kvp.Key] != kvp.Value);
        
        public Dictionary<string, string> GetModifiedValues()
        {
            return _originalValues
                .Where(kvp => _values.ContainsKey(kvp.Key) && _values[kvp.Key] != kvp.Value)
                .ToDictionary(kvp => kvp.Key, kvp => _values[kvp.Key]);
        }

        public void AcceptChanges()
        {
            foreach (var key in _values.Keys.ToList())
            {
                _originalValues[key] = _values[key];
            }
            OnPropertyChanged(nameof(IsModified));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
