

using System.Xml.Serialization;

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
[System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
public partial class cfdi
{

    private cfdiEmisor emisorField;

    private cfdiReceptor receptorField;

    private cfdiConcepto[] conceptosField;

    private string loadField;

    private bool loadFieldSpecified;

    private string tipoDocumentoField;

    //private string RFCEmisorField;

    //private string DireccionEmisorField;

    //private string CPEmisorField;

    //private string RFCReceptorField;

    //private string DireccionReceptorField;

    //private string CPReceptorField;

    private string subTotalField;

    private string totalField;

    private string monedaField;

    private string hazmatField;

    private string tipoContenedorField;

    private string cantBultoField;

    private string pesoTaraField;

    //private bool pesoTaraFieldSpecified;

    private short pesoNetoField;

    private bool pesoNetoFieldSpecified;

    private short pesoBrutoField;

    private bool pesoBrutoFieldSpecified;

    /// <remarks/>
    public cfdiEmisor Emisor
    {
        get
        {
            return this.emisorField;
        }
        set
        {
            this.emisorField = value;
        }
    }

    /// <remarks/>
    public cfdiReceptor Receptor
    {
        get
        {
            return this.receptorField;
        }
        set
        {
            this.receptorField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlArrayItemAttribute("concepto", IsNullable = false)]
    public cfdiConcepto[] Conceptos
    {
        get
        {
            return this.conceptosField;
        }
        set
        {
            this.conceptosField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Load
    {
        get
        {
            return this.loadField;
        }
        set
        {
            this.loadField = value;
        }
    }

    [System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Emisor_Rfc
    //{
    //    get
    //    {
    //        return this.RFCEmisorField;
    //    }
    //    set
    //    {
    //        this.RFCEmisorField = value;
    //    }
    //}

    //[System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Emisor_Direccion
    //{
    //    get
    //    {
    //        return this.DireccionEmisorField;
    //    }
    //    set
    //    {
    //        this.DireccionEmisorField = value;
    //    }
    //}

    //[System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Emisor_codigoPostal
    //{
    //    get
    //    {
    //        return this.CPEmisorField;
    //    }
    //    set
    //    {
    //        this.CPEmisorField = value;
    //    }
    //}

    //[System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Receptor_Rfc
    //{
    //    get
    //    {
    //        return this.RFCReceptorField;
    //    }
    //    set
    //    {
    //        this.RFCReceptorField = value;
    //    }
    //}

    //[System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Receptor_Direccion
    //{
    //    get
    //    {
    //        return this.DireccionReceptorField;
    //    }
    //    set
    //    {
    //        this.DireccionReceptorField = value;
    //    }
    //}

    //[System.Xml.Serialization.XmlAttributeAttribute()]
    //public string Receptor_codigoPostal
    //{
    //    get
    //    {
    //        return this.CPReceptorField;
    //    }
    //    set
    //    {
    //        this.CPReceptorField = value;
    //    }
    //}

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool LoadSpecified
    {
        get
        {
            return this.loadFieldSpecified;
        }
        set
        {
            this.loadFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string TipoDocumento
    {
        get
        {
            return this.tipoDocumentoField;
        }
        set
        {
            this.tipoDocumentoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string SubTotal
    {
        get
        {
            return this.subTotalField;
        }
        set
        {
            this.subTotalField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Total
    {
        get
        {
            return this.totalField;
        }
        set
        {
            this.totalField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Moneda
    {
        get
        {
            return this.monedaField;
        }
        set
        {
            this.monedaField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Hazmat
    {
        get
        {
            return this.hazmatField;
        }
        set
        {
            this.hazmatField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string TipoContenedor
    {
        get
        {
            return this.tipoContenedorField;
        }
        set
        {
            this.tipoContenedorField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CantBulto
    {
        get
        {
            return this.cantBultoField;
        }
        set
        {
            this.cantBultoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string PesoTara
    {
        get
        {
            return this.pesoTaraField;
        }
        set
        {
            this.pesoTaraField = value;
        }
    }

    /// <remarks/>
    //[System.Xml.Serialization.XmlIgnoreAttribute()]
    //public bool PesoTaraSpecified
    //{
    //    get
    //    {
    //        return this.pesoTaraFieldSpecified;
    //    }
    //    set
    //    {
    //        this.pesoTaraFieldSpecified = value;
    //    }
    //}

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public short PesoNeto
    {
        get
        {
            return this.pesoNetoField;
        }
        set
        {
            this.pesoNetoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool PesoNetoSpecified
    {
        get
        {
            return this.pesoNetoFieldSpecified;
        }
        set
        {
            this.pesoNetoFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public short PesoBruto
    {
        get
        {
            return this.pesoBrutoField;
        }
        set
        {
            this.pesoBrutoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool PesoBrutoSpecified
    {
        get
        {
            return this.pesoBrutoFieldSpecified;
        }
        set
        {
            this.pesoBrutoFieldSpecified = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class cfdiEmisor
{

    private string emisor_RfcField;

    private string emisor_NombreField;

    private string emisor_DireccionField;

    private string emisor_codigoPostalField;

    private bool emisor_codigoPostalFieldSpecified;

    private string emisor_paisField;

    private string emisor_estadoField;

    private string emisor_municipioField;

    private string valueField;

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_Rfc
    {
        get
        {
            return this.emisor_RfcField;
        }
        set
        {
            this.emisor_RfcField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_Nombre
    {
        get
        {
            return this.emisor_NombreField;
        }
        set
        {
            this.emisor_NombreField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_Direccion
    {
        get
        {
            return this.emisor_DireccionField;
        }
        set
        {
            this.emisor_DireccionField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_codigoPostal
    {
        get
        {
            return this.emisor_codigoPostalField;
        }
        set
        {
            this.emisor_codigoPostalField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool Emisor_codigoPostalSpecified
    {
        get
        {
            return this.emisor_codigoPostalFieldSpecified;
        }
        set
        {
            this.emisor_codigoPostalFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_pais
    {
        get
        {
            return this.emisor_paisField;
        }
        set
        {
            this.emisor_paisField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_estado
    {
        get
        {
            return this.emisor_estadoField;
        }
        set
        {
            this.emisor_estadoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Emisor_municipio
    {
        get
        {
            return this.emisor_municipioField;
        }
        set
        {
            this.emisor_municipioField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTextAttribute()]
    public string Value
    {
        get
        {
            return this.valueField;
        }
        set
        {
            this.valueField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class cfdiReceptor
{

    private string receptor_RfcField;

    private string receptor_NombreField;

    private string receptor_DireccionField;

    private string receptor_codigoPostalField;

    private bool receptor_codigoPostalFieldSpecified;

    private string receptor_paisField;

    private string receptor_estadoField;

    private string receptor_municipioField;

    private string valueField;

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_Rfc
    {
        get
        {
            return this.receptor_RfcField;
        }
        set
        {
            this.receptor_RfcField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_Nombre
    {
        get
        {
            return this.receptor_NombreField;
        }
        set
        {
            this.receptor_NombreField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_Direccion
    {
        get
        {
            return this.receptor_DireccionField;
        }
        set
        {
            this.receptor_DireccionField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_codigoPostal
    {
        get
        {
            return this.receptor_codigoPostalField;
        }
        set
        {
            this.receptor_codigoPostalField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool Receptor_codigoPostalSpecified
    {
        get
        {
            return this.receptor_codigoPostalFieldSpecified;
        }
        set
        {
            this.receptor_codigoPostalFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_pais
    {
        get
        {
            return this.receptor_paisField;
        }
        set
        {
            this.receptor_paisField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_estado
    {
        get
        {
            return this.receptor_estadoField;
        }
        set
        {
            this.receptor_estadoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string Receptor_municipio
    {
        get
        {
            return this.receptor_municipioField;
        }
        set
        {
            this.receptor_municipioField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTextAttribute()]
    public string Value
    {
        get
        {
            return this.valueField;
        }
        set
        {
            this.valueField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
public partial class cfdiConcepto
{

    private float cantidadField;

    private bool cantidadFieldSpecified;

    private int sAT_ClaveProdServField;

    private bool sAT_ClaveProdServFieldSpecified;

    private string sAT_DescripcionField;

    private string noIdentificacionField;

    private string sAT_UnidadField;

    private string unidadMedidaPesoField;

    private float pesoField;

    private bool pesoFieldSpecified;

    private int facturaField;

    private bool facturaFieldSpecified;

    private long sAT_FraccionArancelariaField;

    private bool sAT_FraccionArancelariaFieldSpecified;

    private string valueField;

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public float Cantidad
    {
        get
        {
            return this.cantidadField;
        }
        set
        {
            this.cantidadField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool CantidadSpecified
    {
        get
        {
            return this.cantidadFieldSpecified;
        }
        set
        {
            this.cantidadFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int SAT_ClaveProdServ
    {
        get
        {
            return this.sAT_ClaveProdServField;
        }
        set
        {
            this.sAT_ClaveProdServField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool SAT_ClaveProdServSpecified
    {
        get
        {
            return this.sAT_ClaveProdServFieldSpecified;
        }
        set
        {
            this.sAT_ClaveProdServFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string SAT_Descripcion
    {
        get
        {
            return this.sAT_DescripcionField;
        }
        set
        {
            this.sAT_DescripcionField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string NoIdentificacion
    {
        get
        {
            return this.noIdentificacionField;
        }
        set
        {
            this.noIdentificacionField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string SAT_Unidad
    {
        get
        {
            return this.sAT_UnidadField;
        }
        set
        {
            this.sAT_UnidadField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string UnidadMedidaPeso
    {
        get
        {
            return this.unidadMedidaPesoField;
        }
        set
        {
            this.unidadMedidaPesoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public float Peso
    {
        get
        {
            return this.pesoField;
        }
        set
        {
            this.pesoField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool PesoSpecified
    {
        get
        {
            return this.pesoFieldSpecified;
        }
        set
        {
            this.pesoFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int Factura
    {
        get
        {
            return this.facturaField;
        }
        set
        {
            this.facturaField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool FacturaSpecified
    {
        get
        {
            return this.facturaFieldSpecified;
        }
        set
        {
            this.facturaFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public long SAT_FraccionArancelaria
    {
        get
        {
            return this.sAT_FraccionArancelariaField;
        }
        set
        {
            this.sAT_FraccionArancelariaField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool SAT_FraccionArancelariaSpecified
    {
        get
        {
            return this.sAT_FraccionArancelariaFieldSpecified;
        }
        set
        {
            this.sAT_FraccionArancelariaFieldSpecified = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTextAttribute()]
    public string Value
    {
        get
        {
            return this.valueField;
        }
        set
        {
            this.valueField = value;
        }
    }
}
