using SolidEdgeCommunity;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace GUI
{
    public partial class Form1 : Form
    {
        private List<Dim> myObjects;
        private List<Dim> selectedObjects;

        public Form1()
        {
            InitializeComponent();
            InitializeData();
            InitializeDataGridView();
        }

        private void InitializeData()
        {
            myObjects = new List<Dim>();
            OleMessageFilter.Register();
            SolidEdgeFramework.Application application = SolidEdgeUtils.Connect(true);
            if (application.ActiveDocumentType != DocumentTypeConstants.igPartDocument) throw new Exception("Current file is not .par!!!");

            var activeDocument = (PartDocument)application.ActiveDocument;

            foreach (ProfileSet profile in activeDocument.ProfileSets)
            {
                foreach (Dimension dim in (Dimensions)profile.Profiles.Item(1).Dimensions)
                {
                    if (dim.IsReadOnly) continue;
                    myObjects.Add(new Dim { Name = dim.VariableTableName, Value = dim.Value, Dimension = dim });
                } 
            }
            OleMessageFilter.Unregister();
        }

        private void InitializeDataGridView()
        {
            dataGridView1.AutoGenerateColumns = false;

            var isSelectedColumn = new DataGridViewCheckBoxColumn
            {
                DataPropertyName = "IsSelected",
                HeaderText = "Select"
            };
            dataGridView1.Columns.Add(isSelectedColumn);

            var nameColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Name",
                HeaderText = "Name"
            };
            dataGridView1.Columns.Add(nameColumn);

            var valueColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Value",
                HeaderText = "Value"
            };
            dataGridView1.Columns.Add(valueColumn);

            dataGridView1.DataSource = myObjects;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == ButtonState.Select.ToString())
            {
                selectedObjects = myObjects.Where(obj => obj.IsSelected).ToList();
                dataGridView1.DataSource = selectedObjects;
                myObjects.ForEach(o => o.IsSelected = false);
                dataGridView1.Columns.RemoveAt(0);
                button2.Text = ButtonState.Save.ToString();
            }
            else if (button2.Text == ButtonState.Save.ToString())
            {
                var isSelectedColumn = new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = "IsSelected",
                    HeaderText = "Select"
                };
                dataGridView1.Columns.Insert(0, isSelectedColumn);
                dataGridView1.DataSource = myObjects;

                foreach (var item in selectedObjects)
                {
                    try
                    {
                        item.Dimension.VariableTableName = item.Name;
                    }
                    catch { }
                    try
                    {
                        item.Dimension.Value = item.Value;
                    }
                    catch { }
                }

                button2.Name = ButtonState.Select.ToString();
            }
        }

        private enum ButtonState
        {
            Select,
            Save
        }
    }
} 
