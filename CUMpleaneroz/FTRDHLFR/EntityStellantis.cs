using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FTRDHLFR
{
    class EntityStellantis
    {
        public string UniqueRecordIdentifier { get; set; }
        public string SenderID {get; set;}
        public string ReciverID {get; set;}
        public string MerchandiseOwnerTaxID {get; set;}
        public string MerchandiseOwner {get; set;}
        public string MerchandiseOwnerAddressLine1 {get; set;}
        public string MerchandiseOwnerAddressLine2 {get; set;}
        public string MerchandiseOwnerCity {get; set;}
        public string MerchandiseOwnerState {get; set;}
        public string MerchandiseOwnerZip {get; set;}
        public string MerchandiseMunicipalySATreference {get; set;}
        public string MerchandiseOwnerCntry {get; set;}
        public string ShipTo {get; set;}
        public string ShipToName {get; set;}
        public string ShipToAddressLine {get; set;}
        public string ShipToCity {get; set;}
        public string ShipToState {get; set;}
        public string ShipToZip {get; set;}
        public string ShipToMunicipalySATreference {get; set;}
        public string ShipToCntry {get; set;}
        public string ShipFrom {get; set;}
        public string SupplierTaxIdentifier {get; set;}
        public string SupplierName {get; set;}
        public string SupplierAddressLine {get; set;}
        public string SupplierCity {get; set;}
        public string SupplierState {get; set;}
        public string SupplierZip {get; set;}
        public string SupplierMuniciplalitySATreference {get; set;}
        public string SupplierCntry {get; set;}
        public string PartOrContainerID {get; set;}
        public string PartDescription {get; set;}
        public string PartSATCode {get; set;}
        public string PartSATDescription {get; set;}
        public string ShippedQuantity {get; set;}
        public string UnitOfMeasureShipped {get; set;}
        public string UnitOfMeasureSATCode {get; set;}
        public string UnitOfMeasureSATdescription {get; set;}
        public string HazmatFlag {get; set;}
        public string HazmatSATCode {get; set;}
        public string HazmatSATDescription {get; set;}
        public string ContainerIdentifier {get; set;}
        public string I_SAT_CNTNR {get; set;}
        public string ContainerSATDescription {get; set;}
        public string ContainerQty {get; set;}
        public string ContainerTareWeight {get; set;}
        public string NetShipmentWeight {get; set;}
        public string GrossShipmentWeight {get; set;}
        public string UnitOfMeasureWeight {get; set;}
        public string HTSCode {get; set;}
        public string HTSCountryCode {get; set;}
        public string CurrencyCode {get; set;}
        public string SupplierCode {get; set;}
        public string FinalDestinationCode {get; set;}
        public string ShipmentIdentifier {get; set;}
        public string SupplierPackingSlip {get; set;}
        public string SupplierBillOfLoading {get; set;}
        public string FreightConsolidationBillOfLandingNumber {get; set;}
        public string ConsolidationShipmentIdentifier {get; set;}
        public string PoolpointShipfrom {get; set;}
        public string PoolPointShipto {get; set;}
        public string ShipmentDate {get; set;}
        public string ShipmentTime {get; set;}
        public string ShipmentTimestamp {get; set;}
        public string CarrierSCAC {get; set;}
        public string ConveyanceIdentifier {get; set;}
        public string OwnerSCAC {get; set;}
        public string TransportationMode {get; set;}
        public string PartCountInContainer {get; set;}
        public string AETCNumer {get; set;}
        public string LotNumber {get; set;}
        public string ChampsTransactionCode {get; set;}
        public string ChampsPurposeCode {get; set;}
        public string ASNStatus {get; set;}
        public string ShipmentIdentifierCount {get; set;}
        public string ASNCount {get; set;}
        public string MasterBilOfLadinng {get; set;}
	    public string UnitEstimatedCost {get; set;}
        public string FillerForFutureUser {get; set;}
        public string TotalWeightContainer { get; set;}
        public string TotalWeightContainerUnit { get; set;}
        public string GrossShipmentWeightUnit { get; set; }
    }
}
