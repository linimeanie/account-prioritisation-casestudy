// Leadership settings used when no settings file is uploaded. Same values as the n8n workflow.
export const DEFAULT_SETTINGS = [
 {
  "section": "general",
  "key": "as_of_date",
  "value": "2026-10-19"
 },
 {
  "section": "general",
  "key": "lookback_days",
  "value": 90
 },
 {
  "section": "general",
  "key": "half_life_days",
  "value": 30
 },
 {
  "section": "general",
  "key": "recent_touch_days",
  "value": 14
 },
 {
  "section": "general",
  "key": "extra_signal_same_type",
  "value": 0.5
 },
 {
  "section": "general",
  "key": "multi_type_bonus",
  "value": 0.15
 },
 {
  "section": "general",
  "key": "max_reasons",
  "value": 3
 }, {
  "section": "general",
  "key": "recycle_after_days",
  "value": 28
 },
 {
  "section": "general",
  "key": "fill_with_fit",
  "value": 1
 }, {
  "section": "general",
  "key": "language_spillover",
  "value": 1
 },
 {
  "section": "signal_weight",
  "key": "Current provider contract renews in ~4 months",
  "value": 10
 },
 {
  "section": "signal_weight",
  "key": "Reported dissatisfaction with coverage on review site",
  "value": 9
 },
 {
  "section": "signal_weight",
  "key": "Contract renewal expected Q1 2027",
  "value": 8
 },
 {
  "section": "signal_weight",
  "key": "3 visits to pricing page",
  "value": 8
 },
 {
  "section": "signal_weight",
  "key": "Visited eSIM product page and pricing",
  "value": 8
 },
 {
  "section": "signal_weight",
  "key": "Visited API docs (5 pages)",
  "value": 7
 },
 {
  "section": "signal_weight",
  "key": "Visited pricing page",
  "value": 6
 },
 {
  "section": "signal_weight",
  "key": "Single visit to blog",
  "value": 1
 },
 {
  "section": "signal_weight",
  "key": "Hiring IoT Connectivity Engineer",
  "value": 7
 },
 {
  "section": "signal_weight",
  "key": "Hiring Firmware Engineer (cellular)",
  "value": 7
 },
 {
  "section": "signal_weight",
  "key": "Hiring Device Operations Manager",
  "value": 5
 },
 {
  "section": "signal_weight",
  "key": "Hiring Embedded Systems Lead",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "Hiring Head of Hardware",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "Downloaded IoT connectivity buyer's guide",
  "value": 6
 },
 {
  "section": "signal_weight",
  "key": "Downloaded eSIM whitepaper",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "Registered for webinar on global roaming",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "Announced next-gen tracker with LTE-M",
  "value": 6
 },
 {
  "section": "signal_weight",
  "key": "Launched new connected device line",
  "value": 5
 },
 {
  "section": "signal_weight",
  "key": "Released fleet platform v3",
  "value": 3
 },
 {
  "section": "signal_weight",
  "key": "New Head of Procurement joined",
  "value": 5
 },
 {
  "section": "signal_weight",
  "key": "New CTO appointed",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "New VP Engineering joined",
  "value": 4
 },
 {
  "section": "signal_weight",
  "key": "New COO from a large OEM",
  "value": 3
 },
 {
  "section": "signal_type_default",
  "key": "Competitor contract renewal due",
  "value": 8
 },
 {
  "section": "signal_type_default",
  "key": "Website visit (pricing or docs page)",
  "value": 5
 },
 {
  "section": "signal_type_default",
  "key": "IoT or hardware job posting",
  "value": 5
 },
 {
  "section": "signal_type_default",
  "key": "Content download",
  "value": 4
 },
 {
  "section": "signal_type_default",
  "key": "Product launch",
  "value": 4
 },
 {
  "section": "signal_type_default",
  "key": "Leadership hire",
  "value": 4
 },
 {
  "section": "signal_type_default",
  "key": "New market entry",
  "value": 5
 },
 {
  "section": "signal_type_default",
  "key": "Funding round",
  "value": 5
 },
 {
  "section": "crm_stage",
  "key": "Never contacted",
  "value": 1.0
 },
 {
  "section": "crm_stage",
  "key": "Contacted",
  "value": 1.0
 },
 {
  "section": "crm_stage",
  "key": "Closed lost",
  "value": 0.8
 },
 {
  "section": "crm_stage",
  "key": "Open opportunity",
  "value": 0
 },
 {
  "section": "crm_stage",
  "key": "Customer",
  "value": 0
 },
 {
  "section": "fit_devices",
  "key": "100000",
  "value": 1.5
 },
 {
  "section": "fit_devices",
  "key": "25000",
  "value": 1.3
 },
 {
  "section": "fit_devices",
  "key": "5000",
  "value": 1.15
 },
 {
  "section": "fit_devices",
  "key": "2000",
  "value": 1.0
 },
 {
  "section": "fit_devices",
  "key": "0",
  "value": 0.85
 },
 {
  "section": "fit_provider",
  "key": "Telvora IoT",
  "value": 1.15
 },
 {
  "section": "fit_provider",
  "key": "Axion Wireless",
  "value": 1.15
 },
 {
  "section": "fit_provider",
  "key": "NordLink M2M",
  "value": 1.15
 },
 {
  "section": "fit_provider",
  "key": "Globecell Connect",
  "value": 1.15
 },
 {
  "section": "fit_provider",
  "key": "Multiple providers",
  "value": 1.15
 },
 {
  "section": "fit_provider",
  "key": "Local MNO",
  "value": 1.1
 },
 {
  "section": "fit_provider",
  "key": "emnify",
  "value": 0
 },
 {
  "section": "fit_size",
  "key": "1-10",
  "value": 0.85
 },
 {
  "section": "fit_size",
  "key": "11-50",
  "value": 1.0
 },
 {
  "section": "fit_size",
  "key": "51-200",
  "value": 1.0
 },
 {
  "section": "fit_size",
  "key": "201-500",
  "value": 1.0
 },
 {
  "section": "fit_size",
  "key": "501-1000",
  "value": 1.05
 },
 {
  "section": "fit_size",
  "key": "1001-5000",
  "value": 1.1
 },
 {
  "section": "fit_size",
  "key": "5000+",
  "value": 1.1
 },
 {
  "section": "boost",
  "key": "industry=Fleet telematics",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Asset tracking",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Cold chain logistics",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Smart metering",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=EV charging",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Micromobility",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Medical devices",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Wearables",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Industrial automation",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Smart buildings",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Security & alarms",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "industry=Agritech",
  "value": 1.0
 },
 {
  "section": "boost",
  "key": "region=EMEA",
  "value": 1.0
 },
 {
  "section": "territory",
  "key": "EMEA",
  "value": "SDR-01"
 },
 {
  "section": "territory",
  "key": "North America",
  "value": "SDR-02,SDR-03"
 },
 {
  "section": "territory",
  "key": "LATAM",
  "value": "SDR-04"
 },
 {
  "section": "territory",
  "key": "APAC",
  "value": ""
 },
 {
  "section": "capacity",
  "key": "SDR-01",
  "value": ""
 },
 {
  "section": "exclude",
  "key": "ACC-0000",
  "value": "example"
 },
 {
  "section": "pin",
  "key": "ACC-0000",
  "value": "example"
 },
 {
  "section": "angle",
  "key": "Competitor contract renewal due",
  "value": "Renewal window: lead with coverage (540+ networks), no lock-in and migration support."
 },
 {
  "section": "angle",
  "key": "Website visit (pricing or docs page)",
  "value": "Active evaluation: offer a pricing walkthrough or free SIM test kit."
 },
 {
  "section": "angle",
  "key": "IoT or hardware job posting",
  "value": "Building connectivity in-house: pitch the API/platform as a shortcut for the new hire."
 },
 {
  "section": "angle",
  "key": "Content download",
  "value": "Research stage: follow up on the asset with a relevant customer story."
 },
 {
  "section": "angle",
  "key": "Product launch",
  "value": "New devices shipping: talk global eSIM provisioning for the launch."
 },
 {
  "section": "angle",
  "key": "Leadership hire",
  "value": "New decision maker: intro + 90-day connectivity review offer."
 },
 {
  "section": "angle",
  "key": "New market entry",
  "value": "Expanding countries: one SIM, local coverage in the new market."
 },
 {
  "section": "angle",
  "key": "Funding round",
  "value": "Fresh budget and scale plans: position for growth and cost control."
 },
 {
  "section": "angle",
  "key": "prospect",
  "value": "No signal yet: research the account, then a cold first touch built on device volume and current provider."
 }
];
