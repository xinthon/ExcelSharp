using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp.Sample
{
    public partial class ViewModel : ObservableObject
    {
        private readonly ExcelSharp excelSharp;

        private readonly ObservableCollection<Model> modelCollection;

        public ObservableCollection<Model> ModelCollection => modelCollection;

        public ViewModel()
        {
            excelSharp = new ExcelSharp();
            modelCollection = new ObservableCollection<Model>();
        }

        [RelayCommand]
        void Import()
        {
            if(FileDialog.Open.ShowDialog() is not true)
                return;

            var filePath = FileDialog.Open.FileName;
            var data = this.excelSharp.Import<Model>(filePath);

            this.ModelCollection.Clear();
            data.ForEach(item => this.ModelCollection.Add(item));
        }

        [RelayCommand]
        void Export()
        {
            if(FileDialog.Save.ShowDialog() is not true)
                return;

            var filePath = FileDialog.Save.FileName;

            var data = new List<Model>();
            this.ModelCollection.ForEach(item => data.Add(item));
            excelSharp.Export(data, filePath);
        }
    }
}
