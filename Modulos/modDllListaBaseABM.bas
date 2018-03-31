Attribute VB_Name = "modDllListaBaseABM"
Option Explicit

'definición de forms y opbjetos a utilizar en ABMs

'forms
Public vFormClPr As frmCListaBaseABM
Public vFormPais As frmCListaBaseABM
Public vFormSucursal As frmCListaBaseABM
Public vFormProvincia As frmCListaBaseABM
Public vFormLocalidad As frmCListaBaseABM
Public vFormMarcas As frmCListaBaseABM
Public vFormTipoComprobante As frmCListaBaseABM
Public vFormCondicionIva As frmCListaBaseABM
Public vFormEstadoDocumento As frmCListaBaseABM
Public vFormFormaPago As frmCListaBaseABM
Public vFormClientes As frmCListaBaseABM
Public vFormVendedor As frmCListaBaseABM
Public vFormProveedor As frmCListaBaseABM
Public vFormTipoProveedor As frmCListaBaseABM
Public vFormTipoGastos As frmCListaBaseABM
Public vFormLineas As frmCListaBaseABM
Public vFormRubros As frmCListaBaseABM
Public vFormProductos As frmCListaBaseABM
Public vFormEstadoProducto As frmCListaBaseABM
Public vFormTipoRevelado As frmCListaBaseABM
Public vFormDestinos As frmCListaBaseABM
Public vFormTarjeta As frmCListaBaseABM
Public vFormTarjetaPlan As frmCListaBaseABM
Public vFormAparato As frmCListaBaseABM
Public vFormObraSocial As frmCListaBaseABM
Public vFormTratamiento As frmCListaBaseABM
Public vFormLabDentales As frmCListaBaseABM
Public vFormLabClinicos As frmCListaBaseABM
Public vFormGrupos As frmCListaBaseABM
Public vFormMedicamentos As frmCListaBaseABM
Public vFormProfesiones As frmCListaBaseABM

'objetos
Public vABMClPr As CListaBaseABM
Public vABMPais As CListaBaseABM
Public vABMSucursal As CListaBaseABM
Public vABMProvincia As CListaBaseABM
Public vABMLocalidad As CListaBaseABM
Public vABMMarcas As CListaBaseABM
Public vABMTipoCompronate As CListaBaseABM
Public vABMCondicionIva As CListaBaseABM
Public vABMEstadoDocumento As CListaBaseABM
Public vABMFormaPago As CListaBaseABM
Public vABMClientes As CListaBaseABM
Public vABMVendedor As CListaBaseABM
Public vABMProveedor As CListaBaseABM
Public vABMTipoProveedor As CListaBaseABM
Public vABMTipoGastos As CListaBaseABM
Public vABMLineas As CListaBaseABM
Public vABMRubros As CListaBaseABM
Public vABMProductos As CListaBaseABM
Public vABMEstadoProducto As CListaBaseABM
Public vABMTipoRevelado As CListaBaseABM
Public vABMDestinos As CListaBaseABM
Public vABMTarjeta As CListaBaseABM
Public vABMTarjetaPlan As CListaBaseABM
Public vABMAparato As CListaBaseABM
Public vABMObraSocial As CListaBaseABM
Public vABMTratamiento As CListaBaseABM
Public vABMLabDentales As CListaBaseABM
Public vABMLabClinicos As CListaBaseABM
Public vABMGrupos As CListaBaseABM
Public vABMMedicamentos As CListaBaseABM
Public vABMProfesiones As CListaBaseABM

'variable para mantener el objeto base de ABM activo
Public auxDllActiva As CListaBaseABM
'Public auxDllActivaCta As CListaBaseABMCta

