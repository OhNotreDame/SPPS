(function () {

    //Initialize the variables for overrides objects
    var overrideCtx = {};
    overrideCtx.Templates = {};
	overrideCtx.ListTemplateType = 103; //Links list only
    overrideCtx.Templates.Fields = {
        'URLwMenu': { 'View': getLinkLocation },
        'URLNoMenu': { 'View': getLinkLocation },
        'URL': { 'View': getLinkLocation }
    };

    //Register the template overrides.
    SPClientTemplates.TemplateManager.RegisterTemplateOverrides(overrideCtx);

})();

function getLinkLocation(ctx) {

	//--------------------------------------//
	//----------- OVERRIDE UTILE -----------//
    //--------------------------------------//
    return '<a href=\"' + ctx.CurrentItem.URL + '\" target=\"_blank\">' + ctx.CurrentItem['URL.desc'] + '</a>'
}