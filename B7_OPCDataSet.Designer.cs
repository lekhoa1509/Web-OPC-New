﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 1591

namespace web4 {
    
    
    /// <summary>
    ///Represents a strongly typed in-memory cache of data.
    ///</summary>
    [global::System.Serializable()]
    [global::System.ComponentModel.DesignerCategoryAttribute("code")]
    [global::System.ComponentModel.ToolboxItem(true)]
    [global::System.Xml.Serialization.XmlSchemaProviderAttribute("GetTypedDataSetSchema")]
    [global::System.Xml.Serialization.XmlRootAttribute("B7_OPCDataSet")]
    [global::System.ComponentModel.Design.HelpKeywordAttribute("vs.data.DataSet")]
    public partial class B7_OPCDataSet : global::System.Data.DataSet {
        
        private global::System.Data.SchemaSerializationMode _schemaSerializationMode = global::System.Data.SchemaSerializationMode.IncludeSchema;
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        public B7_OPCDataSet() {
            this.BeginInit();
            this.InitClass();
            global::System.ComponentModel.CollectionChangeEventHandler schemaChangedHandler = new global::System.ComponentModel.CollectionChangeEventHandler(this.SchemaChanged);
            base.Tables.CollectionChanged += schemaChangedHandler;
            base.Relations.CollectionChanged += schemaChangedHandler;
            this.EndInit();
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected B7_OPCDataSet(global::System.Runtime.Serialization.SerializationInfo info, global::System.Runtime.Serialization.StreamingContext context) : 
                base(info, context, false) {
            if ((this.IsBinarySerialized(info, context) == true)) {
                this.InitVars(false);
                global::System.ComponentModel.CollectionChangeEventHandler schemaChangedHandler1 = new global::System.ComponentModel.CollectionChangeEventHandler(this.SchemaChanged);
                this.Tables.CollectionChanged += schemaChangedHandler1;
                this.Relations.CollectionChanged += schemaChangedHandler1;
                return;
            }
            string strSchema = ((string)(info.GetValue("XmlSchema", typeof(string))));
            if ((this.DetermineSchemaSerializationMode(info, context) == global::System.Data.SchemaSerializationMode.IncludeSchema)) {
                global::System.Data.DataSet ds = new global::System.Data.DataSet();
                ds.ReadXmlSchema(new global::System.Xml.XmlTextReader(new global::System.IO.StringReader(strSchema)));
                this.DataSetName = ds.DataSetName;
                this.Prefix = ds.Prefix;
                this.Namespace = ds.Namespace;
                this.Locale = ds.Locale;
                this.CaseSensitive = ds.CaseSensitive;
                this.EnforceConstraints = ds.EnforceConstraints;
                this.Merge(ds, false, global::System.Data.MissingSchemaAction.Add);
                this.InitVars();
            }
            else {
                this.ReadXmlSchema(new global::System.Xml.XmlTextReader(new global::System.IO.StringReader(strSchema)));
            }
            this.GetSerializationData(info, context);
            global::System.ComponentModel.CollectionChangeEventHandler schemaChangedHandler = new global::System.ComponentModel.CollectionChangeEventHandler(this.SchemaChanged);
            base.Tables.CollectionChanged += schemaChangedHandler;
            this.Relations.CollectionChanged += schemaChangedHandler;
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        [global::System.ComponentModel.BrowsableAttribute(true)]
        [global::System.ComponentModel.DesignerSerializationVisibilityAttribute(global::System.ComponentModel.DesignerSerializationVisibility.Visible)]
        public override global::System.Data.SchemaSerializationMode SchemaSerializationMode {
            get {
                return this._schemaSerializationMode;
            }
            set {
                this._schemaSerializationMode = value;
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        [global::System.ComponentModel.DesignerSerializationVisibilityAttribute(global::System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public new global::System.Data.DataTableCollection Tables {
            get {
                return base.Tables;
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        [global::System.ComponentModel.DesignerSerializationVisibilityAttribute(global::System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public new global::System.Data.DataRelationCollection Relations {
            get {
                return base.Relations;
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected override void InitializeDerivedDataSet() {
            this.BeginInit();
            this.InitClass();
            this.EndInit();
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        public override global::System.Data.DataSet Clone() {
            B7_OPCDataSet cln = ((B7_OPCDataSet)(base.Clone()));
            cln.InitVars();
            cln.SchemaSerializationMode = this.SchemaSerializationMode;
            return cln;
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected override bool ShouldSerializeTables() {
            return false;
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected override bool ShouldSerializeRelations() {
            return false;
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected override void ReadXmlSerializable(global::System.Xml.XmlReader reader) {
            if ((this.DetermineSchemaSerializationMode(reader) == global::System.Data.SchemaSerializationMode.IncludeSchema)) {
                this.Reset();
                global::System.Data.DataSet ds = new global::System.Data.DataSet();
                ds.ReadXml(reader);
                this.DataSetName = ds.DataSetName;
                this.Prefix = ds.Prefix;
                this.Namespace = ds.Namespace;
                this.Locale = ds.Locale;
                this.CaseSensitive = ds.CaseSensitive;
                this.EnforceConstraints = ds.EnforceConstraints;
                this.Merge(ds, false, global::System.Data.MissingSchemaAction.Add);
                this.InitVars();
            }
            else {
                this.ReadXml(reader);
                this.InitVars();
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected override global::System.Xml.Schema.XmlSchema GetSchemaSerializable() {
            global::System.IO.MemoryStream stream = new global::System.IO.MemoryStream();
            this.WriteXmlSchema(new global::System.Xml.XmlTextWriter(stream, null));
            stream.Position = 0;
            return global::System.Xml.Schema.XmlSchema.Read(new global::System.Xml.XmlTextReader(stream), null);
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        internal void InitVars() {
            this.InitVars(true);
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        internal void InitVars(bool initTable) {
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        private void InitClass() {
            this.DataSetName = "B7_OPCDataSet";
            this.Prefix = "";
            this.Namespace = "http://tempuri.org/B7_OPCDataSet.xsd";
            this.EnforceConstraints = true;
            this.SchemaSerializationMode = global::System.Data.SchemaSerializationMode.IncludeSchema;
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        private void SchemaChanged(object sender, global::System.ComponentModel.CollectionChangeEventArgs e) {
            if ((e.Action == global::System.ComponentModel.CollectionChangeAction.Remove)) {
                this.InitVars();
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        public static global::System.Xml.Schema.XmlSchemaComplexType GetTypedDataSetSchema(global::System.Xml.Schema.XmlSchemaSet xs) {
            B7_OPCDataSet ds = new B7_OPCDataSet();
            global::System.Xml.Schema.XmlSchemaComplexType type = new global::System.Xml.Schema.XmlSchemaComplexType();
            global::System.Xml.Schema.XmlSchemaSequence sequence = new global::System.Xml.Schema.XmlSchemaSequence();
            global::System.Xml.Schema.XmlSchemaAny any = new global::System.Xml.Schema.XmlSchemaAny();
            any.Namespace = ds.Namespace;
            sequence.Items.Add(any);
            type.Particle = sequence;
            global::System.Xml.Schema.XmlSchema dsSchema = ds.GetSchemaSerializable();
            if (xs.Contains(dsSchema.TargetNamespace)) {
                global::System.IO.MemoryStream s1 = new global::System.IO.MemoryStream();
                global::System.IO.MemoryStream s2 = new global::System.IO.MemoryStream();
                try {
                    global::System.Xml.Schema.XmlSchema schema = null;
                    dsSchema.Write(s1);
                    for (global::System.Collections.IEnumerator schemas = xs.Schemas(dsSchema.TargetNamespace).GetEnumerator(); schemas.MoveNext(); ) {
                        schema = ((global::System.Xml.Schema.XmlSchema)(schemas.Current));
                        s2.SetLength(0);
                        schema.Write(s2);
                        if ((s1.Length == s2.Length)) {
                            s1.Position = 0;
                            s2.Position = 0;
                            for (; ((s1.Position != s1.Length) 
                                        && (s1.ReadByte() == s2.ReadByte())); ) {
                                ;
                            }
                            if ((s1.Position == s1.Length)) {
                                return type;
                            }
                        }
                    }
                }
                finally {
                    if ((s1 != null)) {
                        s1.Close();
                    }
                    if ((s2 != null)) {
                        s2.Close();
                    }
                }
            }
            xs.Add(dsSchema);
            return type;
        }
    }
}
namespace web4.B7_OPCDataSetTableAdapters {
    
    
    /// <summary>
    ///Represents the connection and commands used to retrieve and save data.
    ///</summary>
    [global::System.ComponentModel.DesignerCategoryAttribute("code")]
    [global::System.ComponentModel.ToolboxItem(true)]
    [global::System.ComponentModel.DataObjectAttribute(true)]
    [global::System.ComponentModel.DesignerAttribute("Microsoft.VSDesigner.DataSource.Design.TableAdapterDesigner, Microsoft.VSDesigner" +
        ", Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
    [global::System.ComponentModel.Design.HelpKeywordAttribute("vs.data.TableAdapter")]
    public partial class QueriesTableAdapter : global::System.ComponentModel.Component {
        
        private global::System.Data.IDbCommand[] _commandCollection;
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        protected global::System.Data.IDbCommand[] CommandCollection {
            get {
                if ((this._commandCollection == null)) {
                    this.InitCommandCollection();
                }
                return this._commandCollection;
            }
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        private void InitCommandCollection() {
            this._commandCollection = new global::System.Data.IDbCommand[1];
            this._commandCollection[0] = new global::System.Data.SqlClient.SqlCommand();
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Connection = new global::System.Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["Model1"].ConnectionString);
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).CommandText = "dbo.usp_Vth_BC_BHCNTK_CN";
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).CommandType = global::System.Data.CommandType.StoredProcedure;
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@RETURN_VALUE", global::System.Data.SqlDbType.Int, 4, global::System.Data.ParameterDirection.ReturnValue, 10, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ngay_Ct1", global::System.Data.SqlDbType.SmallDateTime, 4, global::System.Data.ParameterDirection.Input, 16, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ngay_Ct2", global::System.Data.SqlDbType.SmallDateTime, 4, global::System.Data.ParameterDirection.Input, 16, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Kho", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Bp", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_TDV", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Vt", global::System.Data.SqlDbType.NVarChar, 24, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Vt_Khac_List", global::System.Data.SqlDbType.NVarChar, 192, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Dt", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_LKH", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_LSP", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Vung", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_BizDocId_C2", global::System.Data.SqlDbType.NVarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_CtLst", global::System.Data.SqlDbType.NVarChar, 128, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Khuyen_Mai", global::System.Data.SqlDbType.Bit, 1, global::System.Data.ParameterDirection.Input, 1, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_LangId", global::System.Data.SqlDbType.Int, 4, global::System.Data.ParameterDirection.Input, 10, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Dvcs", global::System.Data.SqlDbType.NChar, 3, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_Dvcs1", global::System.Data.SqlDbType.NChar, 254, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_User_Name", global::System.Data.SqlDbType.VarChar, 16, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_Ma_TTe0", global::System.Data.SqlDbType.NVarChar, 3, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_BatchCode", global::System.Data.SqlDbType.NVarChar, 24, global::System.Data.ParameterDirection.Input, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
            ((global::System.Data.SqlClient.SqlCommand)(this._commandCollection[0])).Parameters.Add(new global::System.Data.SqlClient.SqlParameter("@_LAYOUT_XML", global::System.Data.SqlDbType.NVarChar, 2147483647, global::System.Data.ParameterDirection.InputOutput, 0, 0, null, global::System.Data.DataRowVersion.Current, false, null, "", "", ""));
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "4.0.0.0")]
        [global::System.ComponentModel.Design.HelpKeywordAttribute("vs.data.TableAdapter")]
        public virtual int usp_Vth_BC_BHCNTK_CN(
                    global::System.Nullable<global::System.DateTime> _Ngay_Ct1, 
                    global::System.Nullable<global::System.DateTime> _Ngay_Ct2, 
                    string _Ma_Kho, 
                    string _Ma_Bp, 
                    string _Ma_TDV, 
                    string _Ma_Vt, 
                    string _Ma_Vt_Khac_List, 
                    string _Ma_Dt, 
                    string _Ma_LKH, 
                    string _Ma_LSP, 
                    string _Ma_Vung, 
                    string _BizDocId_C2, 
                    string _Ma_CtLst, 
                    global::System.Nullable<bool> _Khuyen_Mai, 
                    global::System.Nullable<int> _LangId, 
                    string _Ma_Dvcs, 
                    string _Ma_Dvcs1, 
                    string _User_Name, 
                    string _Ma_TTe0, 
                    string _BatchCode, 
                    ref string _LAYOUT_XML) {
            global::System.Data.SqlClient.SqlCommand command = ((global::System.Data.SqlClient.SqlCommand)(this.CommandCollection[0]));
            if ((_Ngay_Ct1.HasValue == true)) {
                command.Parameters[1].Value = ((System.DateTime)(_Ngay_Ct1.Value));
            }
            else {
                command.Parameters[1].Value = global::System.DBNull.Value;
            }
            if ((_Ngay_Ct2.HasValue == true)) {
                command.Parameters[2].Value = ((System.DateTime)(_Ngay_Ct2.Value));
            }
            else {
                command.Parameters[2].Value = global::System.DBNull.Value;
            }
            if ((_Ma_Kho == null)) {
                command.Parameters[3].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[3].Value = ((string)(_Ma_Kho));
            }
            if ((_Ma_Bp == null)) {
                command.Parameters[4].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[4].Value = ((string)(_Ma_Bp));
            }
            if ((_Ma_TDV == null)) {
                command.Parameters[5].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[5].Value = ((string)(_Ma_TDV));
            }
            if ((_Ma_Vt == null)) {
                command.Parameters[6].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[6].Value = ((string)(_Ma_Vt));
            }
            if ((_Ma_Vt_Khac_List == null)) {
                command.Parameters[7].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[7].Value = ((string)(_Ma_Vt_Khac_List));
            }
            if ((_Ma_Dt == null)) {
                command.Parameters[8].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[8].Value = ((string)(_Ma_Dt));
            }
            if ((_Ma_LKH == null)) {
                command.Parameters[9].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[9].Value = ((string)(_Ma_LKH));
            }
            if ((_Ma_LSP == null)) {
                command.Parameters[10].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[10].Value = ((string)(_Ma_LSP));
            }
            if ((_Ma_Vung == null)) {
                command.Parameters[11].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[11].Value = ((string)(_Ma_Vung));
            }
            if ((_BizDocId_C2 == null)) {
                command.Parameters[12].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[12].Value = ((string)(_BizDocId_C2));
            }
            if ((_Ma_CtLst == null)) {
                command.Parameters[13].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[13].Value = ((string)(_Ma_CtLst));
            }
            if ((_Khuyen_Mai.HasValue == true)) {
                command.Parameters[14].Value = ((bool)(_Khuyen_Mai.Value));
            }
            else {
                command.Parameters[14].Value = global::System.DBNull.Value;
            }
            if ((_LangId.HasValue == true)) {
                command.Parameters[15].Value = ((int)(_LangId.Value));
            }
            else {
                command.Parameters[15].Value = global::System.DBNull.Value;
            }
            if ((_Ma_Dvcs == null)) {
                command.Parameters[16].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[16].Value = ((string)(_Ma_Dvcs));
            }
            if ((_Ma_Dvcs1 == null)) {
                command.Parameters[17].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[17].Value = ((string)(_Ma_Dvcs1));
            }
            if ((_User_Name == null)) {
                command.Parameters[18].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[18].Value = ((string)(_User_Name));
            }
            if ((_Ma_TTe0 == null)) {
                command.Parameters[19].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[19].Value = ((string)(_Ma_TTe0));
            }
            if ((_BatchCode == null)) {
                command.Parameters[20].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[20].Value = ((string)(_BatchCode));
            }
            if ((_LAYOUT_XML == null)) {
                command.Parameters[21].Value = global::System.DBNull.Value;
            }
            else {
                command.Parameters[21].Value = ((string)(_LAYOUT_XML));
            }
            global::System.Data.ConnectionState previousConnectionState = command.Connection.State;
            if (((command.Connection.State & global::System.Data.ConnectionState.Open) 
                        != global::System.Data.ConnectionState.Open)) {
                command.Connection.Open();
            }
            int returnValue;
            try {
                returnValue = command.ExecuteNonQuery();
            }
            finally {
                if ((previousConnectionState == global::System.Data.ConnectionState.Closed)) {
                    command.Connection.Close();
                }
            }
            if (((command.Parameters[21].Value == null) 
                        || (command.Parameters[21].Value.GetType() == typeof(global::System.DBNull)))) {
                _LAYOUT_XML = null;
            }
            else {
                _LAYOUT_XML = ((string)(command.Parameters[21].Value));
            }
            return returnValue;
        }
    }
}

#pragma warning restore 1591