import { IInputs, IOutputs } from "./generated/ManifestTypes";

//https://stackoverflow.com/questions/13689705/how-to-add-google-maps-autocomplete-search-box for geolocating
//https://azuremapscodesamples.azurewebsites.net/ for samples
//https://github.com/Azure-Samples/AzureMapsCodeSamples/blob/master/AzureMapsCodeSamples/REST%20Services/Search%20Autosuggest%20and%20JQuery%20UI.html

export class AzureMapsAddressAutoCompletePCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private notifyOutputChanged: () => void;
    private searchBox: HTMLInputElement;
    private btnEdit: HTMLAnchorElement;
    private checkBoxGlobalSearch: HTMLInputElement;
    private divSearch: HTMLDivElement;

    private value: string;
    private street: string;
    private city: string;
    private county: string;
    private state: string;
    private zipcode: string;
    private country: string;
    private latitude: string;
    private longitude: string

    constructor() {

    }

    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {

        let scriptUrl = "https://atlas.microsoft.com/sdk/javascript/mapcontrol/2/atlas.min.js";
        let scriptNode = document.createElement("script");
        scriptNode.setAttribute("type", "text/javascript");
        scriptNode.setAttribute("src", scriptUrl);
        document.head.appendChild(scriptNode);

        /* used in testing */
        // let jqueryScriptUrl = 'https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js';
        // let jqueryScript = document.createElement("script");
        // jqueryScript.setAttribute("type", "text/javascript");
        // jqueryScript.setAttribute("src", jqueryScriptUrl);
        // document.head.appendChild(jqueryScript);

        let jqueryUIScriptUrl = 'https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js';
        let jqueryUIScript = document.createElement("script");
        jqueryUIScript.setAttribute("type", "text/javascript");
        jqueryUIScript.setAttribute("src", jqueryUIScriptUrl);
        document.head.appendChild(jqueryUIScript);
    
        let jqueryUISmoothLinkCSS = 'https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css';
        let jqueryUISmoothlink = document.createElement("link");
        jqueryUISmoothlink.setAttribute("rel", "stylesheet");
        jqueryUISmoothlink.setAttribute("href", jqueryUISmoothLinkCSS);
        document.head.appendChild(jqueryUISmoothlink);
        
        this.notifyOutputChanged = notifyOutputChanged;     
        let addresssGeocodeServiceUrlTemplate = 'https://atlas.microsoft.com/search/address/json?typeahead=true&subscription-key={subscription-key}&api-version=1&query={query}&language={language}&countrySet={countrySet}&view=Auto';
        let subscriptionKey  : string = context.parameters.azuremapsapikey.raw;
        let streetPopulated : boolean = context.parameters.street.raw !== undefined && context.parameters.street.raw !== null && context.parameters.street.raw !== "";
        

        /* https://docs.microsoft.com/en-us/azure/azure-maps/how-to-manage-account-keys */
        /*  */
        // let subscriptionKey : string = "<Your subscription key for testing>";
        // let currentUserId : string = "<Your User Id for testing>";
        // let currentClientURL : string = "<Your client URL for testing>";


        /* Filter by countries section */
        //let currentUserId : string = Xrm.Page.context.getUserId().replace("{","").replace("}","");
        //let currentClientURL : string = Xrm.Page.context.getClientUrl();
        // "Created a field on a custom entity called 'Country' which was added to the 'User' record.  Used the CRM https://github.com/jlattimer/CRMRESTBuilder to help build fetch xml to pull a comma delimited list of countries to filter current users search results
        // Format something like let countrySearchFetchXML = currentClientURL + "BUILD FETCH XML LINE HERE USING CURRENT USER AS CONDITION" + currentUserId + "END OF FETCH"
        // Replace line below
        let countrySearchFetchXML = "";

        this.btnEdit = document.createElement("a");
        this.btnEdit.setAttribute("href", "javascript:void(0)");
        this.btnEdit.setAttribute("class", "ui-button ui-widget ui-corner-all");
        this.btnEdit.addEventListener("click", this.onEdit.bind(this));
        this.btnEdit.innerText = "Change Address";
        container.appendChild(this.btnEdit);

        this.divSearch = document.createElement("div");
        this.divSearch.setAttribute("class", "widget");
        this.divSearch.setAttribute("style", "text-align:left");
        container.appendChild(this.divSearch);
    

        this.searchBox = document.createElement("input");
        this.searchBox.setAttribute("id", "azureMapSearchBox");
        this.searchBox.setAttribute("style", "width:80%");
        this.searchBox.className = "addressAutocomplete";
        this.searchBox.addEventListener("mouseenter", this.onMouseEnter.bind(this));
        this.searchBox.addEventListener("mouseleave", this.onMouseLeave.bind(this));
        this.divSearch.appendChild(this.searchBox);

        this.checkBoxGlobalSearch = document.createElement("input");
        this.checkBoxGlobalSearch.setAttribute("type", "checkbox");
        this.checkBoxGlobalSearch.setAttribute("name", "checkbox-1");
        this.checkBoxGlobalSearch.setAttribute("id", "checkbox-1");
        this.divSearch.appendChild(this.checkBoxGlobalSearch);

        let labelForCheckBoxGlobalSearch = document.createElement("label");
        labelForCheckBoxGlobalSearch.setAttribute("for", "checkbox-1");
        labelForCheckBoxGlobalSearch.setAttribute("class", "ui-checkboxradio-label ui-widget");
        labelForCheckBoxGlobalSearch.innerText = "Global Search";
        this.divSearch.appendChild(labelForCheckBoxGlobalSearch);

        if(!streetPopulated) {
            this.btnEdit.setAttribute("style", "display: none;");
            this.divSearch.setAttribute("style", "display: inline;");

        }
        else {
            this.btnEdit.setAttribute("style", "display: inline;");
            this.divSearch.setAttribute("style", "display: none;");
        }

        let userCountry = (() => {
        //Add azure map countries to country
            var results: string = "";
            if(this.checkBoxGlobalSearch.checked)
                results = 'AX,AL,DZ,AS,AD,AO,AI,AQ,AG,AR,AM,AW,AU,AT,AZ,BS,BH,BD,BB,BY,BE,BZ,BJ,BM,BT,BO,BQ,BA,BW,BV,BR,IO,BN,BG,BF,BI,CV,KH,CM,CA,KY,CF,TD,CL,CN,CX,CC,CO,KM,CG,CD,CK,CR,CI,HR,CU,CW,CY,CZ,DK,DJ,DM,DO,EC,EG,SV,GQ,ER,EE,SZ,ET,FK,FO,FJ,FI,FR,GF,PF,TF,GA,GM,GE,DE,GH,GI,GR,GL,GD,GP,GU,GT,GG,GN,GW,GY,HT,HM,VA,HN,HK,HU,IS,IN,ID,IR,IQ,IE,IM,IL,IT,JM,JP,JE,JO,KZ,KE,KI,KP,KR,KW,KG,LA,LV,LB,LS,LR,LY,LI,LT,LU,MO,MK,MG,MW,MY,MV,ML,MT,MH,MQ,MR,MU,YT,MX,FM,MD,MC,MN,ME,MS,MA,MZ,MM,NA,NR,NP,NL,NC,NZ,NI,NE,NG,NU,NF,MP,NO,OM,PK,PW,PS,PA,PG,PY,PE,PH,PN,PL,PT,PR,QA,RE,RO,RU,RW,BL,SH,KN,LC,MF,PM,VC,WS,SM,ST,SA,SN,RS,SC,SL,SG,SX,SK,SI,SB,SO,ZA,GS,SS,ES,LK,SD,SR,SJ,SE,CH,SY,TW,TJ,TZ,TH,TL,TG,TK,TO,TT,TN,TR,TM,TC,TV,UG,UA,AE,GB,UM,US,UY,UZ,VU,VE,VN,VG,VI,WF,EH,YE,ZM,ZW';
            else if (countrySearchFetchXML !== "")
            {
                $.ajax({
                    type: "GET",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    url: countrySearchFetchXML,
                    beforeSend: function(XMLHttpRequest) {
                        XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                        XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                        XMLHttpRequest.setRequestHeader("Accept", "application/json");
                        XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                    },
                    async: false,
                    success: function(data : any) {
                        if(data.value !== null && data.value[0] !== null)
                            results = data.value[0]["<comma delimited countries>"];
                    },
                    error: function(xhr, textStatus, errorThrown) {
                        console.log(textStatus + " " + errorThrown);
                    }
                });
           }
            return results === "" ? "US" : results; 
        });

        window.setTimeout(() => {
            $("#azureMapSearchBox").autocomplete({
                minLength: 3,   //Don't ask for suggestions until atleast 3 characters have been typed.
                source: function (request: any, response: any) {
                    //Create a URL to the Azure Maps search service to perform the address search.
                    var requestUrl = addresssGeocodeServiceUrlTemplate.replace('{query}', encodeURIComponent(request.term))
                        .replace('{subscription-key}', subscriptionKey)
                        .replace('{language}', 'en-US')
                        .replace('{countrySet}', userCountry); //A comma seperated string of country codes to limit the suggestions to.
                    $.ajax({
                        url: requestUrl,
                        success: function (data: any) {
                            response(data.results);
                        }
                    });
                },
                select: (event: any, ui: any) => {

                    var selection = ui.item;

                    this.value = "";
                    this.street = (selection.address.streetNumber ? (selection.address.streetNumber + ' ') : '') + (selection.address.streetName || '');
                    this.city = selection.address.municipality || '';
                    this.county = selection.address.countrySecondarySubdivision || '';
                    this.state = selection.address.countrySubdivision || '';
                    this.country = selection.address.countryCodeISO3 || '';
                    this.zipcode = selection.address.postalCode || '';
                    this.latitude = selection.position.lat.toString().trim() || '';
                    this.longitude = selection.position.lon.toString().trim() || '';
                    this.notifyOutputChanged();

                }
            }).data("ui-autocomplete")._renderItem = (ul: any, item: any) => {
                //Format the displayed suggestion to show the formatted suggestion string.
                var suggestionLabel = item.address.freeformAddress;
                if (item.poi && item.poi.name) {
                    suggestionLabel = item.poi.name + ' (' + suggestionLabel + ')';
                }
                return $("<li>")
                    .append("<a>" + suggestionLabel + "</a>")
                    .appendTo(ul);
            };
        },
            2000);
    }

    private onMouseEnter(): void {
        this.searchBox.className = "addressAutocompleteFocused";
    }

    private onMouseLeave(): void {
        this.searchBox.className = "addressAutocomplete";
    }

    private onEdit(): void {
        //this.searchBox.disabled = false;
        this.btnEdit.setAttribute("style", "display: none;");
        this.divSearch.setAttribute("style", "display: inline;");
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
    }

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
    public getOutputs(): IOutputs {
        return {
            value: this.value,
            street: this.street,
            city: this.city,
            county: this.county,
            state: this.state,
            country: this.country,
            zipcode: this.zipcode,
            latitude: this.latitude,
            longitude: this.longitude
        };
    }

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}