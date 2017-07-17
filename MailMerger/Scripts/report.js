
// executes when complete page is fully loaded, including all frames, objects and images
var email, file_type, split, prefix, field1, field2, field3, field4, field5, suffix, delimiter;

$(window).load(function () {


    var delay = 3000;
    setInterval(function () {

        if (top.document.title.indexOf('Mail Merge Report') > -1) {

            var downloadLink = $('.gFiles > div > fieldset > a').attr('href');
            if ($('.gNormalMenu > a[title="Download Report"]').length < 1 && downloadLink != null) {
                setup();
            }

            if ($('.gNormalMenu > a[title="Download Template"]').length < 1) {
                var TemplateFile = $('.gField[size="50"]').val();
                if (TemplateFile != null && TemplateFile != '')
                    mergeSetup();
            }

        } else if (top.document.title.indexOf('Producing Letter/Cert') > -1) {

            var downloadLink = $('.gFiles > div > fieldset > a').attr('href');
            if ($('.gNormalMenu > a[title="Download Report"]').length < 1 && downloadLink != null) {

                setup2();
            }

            if ($('.gNormalMenu > a[title="Download Template"]').length < 1) {
                var TemplateFile = $('.gField[size="50"]').val();
                if (TemplateFile != null && TemplateFile != '')
                    mergeSetup();
            }
        }

    }, delay);
});


function mergeSetup() {

    var dbutton = $('<a class="gClickable gMenuAction gMenuAction4st" title="Download Template" href="#">Download Template</a>');
    dbutton.click(function () {

        var TemplateFile = $('.gField[size="50"]').val();
        if (TemplateFile != null && TemplateFile != '') {

            var win = window.open('http://10.10.10.18/LPMailMergeWeb/CreateTemplate.aspx?path=' + encodeURIComponent(TemplateFile), 'Download', 'width=100,height=100,resizable=no,scrollbars=no,toolbar=no,location=no,status=no');
        }
    });
    $('.gNormalMenu > a[title="Run An Existing Report"]').after(dbutton);
    $('.gToolBar button[title="OK"]').on('click', function () {

        email = $('.gcMM3_email>input').val();
        file_type = $('.gcMM3_gns_file_type input').val();
        split = $('.gcMM3_gns_split input').val();
        prefix = $('.gcMM3_gns_prefix input').val();
        field1 = $('.gcMM3_gns_field1 input').val();
        field2 = $('.gcMM3_gns_field2 input').val();
        field3 = $('.gcMM3_gns_field3 input').val();
        field4 = $('.gcMM3_gns_field4 input').val();
        field5 = $('.gcMM3_gns_field5 input').val();
        suffix = $('.gcMM3_gns_suffix input').val();
        delimiter = $('.gcMM3_gns_delimiter input').val();
    });
}


function setup() {

    var dbutton = $('<a class="gClickable gMenuAction gMenuAction4st" title="Download Report" href="#">Download Merged Report</a>');
    dbutton.click(function () {

        var downloadLink = $('.gFiles > div > fieldset > a').attr('href');

        var TemplateFile = $('.gField[size="50"]')[0].value;

        var win = window.open('http://10.10.10.18/LPMailMergeWeb/MailMerge.aspx?format=' + encodeURIComponent(TemplateFile) + '&email=' + email + '&file_type=' + file_type + '&split=' + split + '&prefix=' + prefix + '&field1=' + field1 + '&field2=' + field2 + '&field3=' + field3 + '&field4=' + field4 + '&field5=' + field5 + '&suffix=' + suffix + '&delimiter=' + delimiter + '&source=' + window.location.protocol
+ "//" + window.location.host + downloadLink, 'Download', 'width=200,height=200,toolbar=no,location=no,status=no');

    });
    $('.gNormalMenu > a[title="Run An Existing Report"]').after(dbutton);

}

function setup2() {


    var dbutton = $('<a class="gClickable gMenuAction gMenuAction4st" title="Download Report" href="#">Download Merged Report</a>');
    dbutton.click(function () {

        var downloadLink = $('.gFiles > div > fieldset > a').attr('href');


        //var TemplateFile = $('.gField[size="50"]')[0].value;

        var TemplateFile = $('.gcML_document>input').val();
        email = $('.gcML_email>input').val();
        file_type = "Word document";
        split = "No";
        //alert(downloadLink);

        //alert(TemplateFile);
        var sourceURL = "";
        if (downloadLink.indexOf("10.10.10.18") == -1) {
            sourceURL = window.location.protocol + "//" + window.location.host + downloadLink;
        } else {
            sourceURL = downloadLink;
        }

        var win = window.open('http://10.10.10.18/LPMailMergeWeb/MailMerge.aspx?format=' + encodeURIComponent(TemplateFile) + '&email=' + email + '&file_type=' + file_type + '&split=' + split + '&prefix=' + prefix + '&field1=' + field1 + '&field2=' + field2 + '&field3=' + field3 + '&field4=' + field4 + '&field5=' + field5 + '&suffix=' + suffix + '&delimiter=' + delimiter + '&source=' + window.location.protocol
    + "//" + window.location.host + downloadLink, 'Download', 'width=200,height=200,toolbar=no,location=no,status=no');


    });

    $('.gNormalMenu > a[title="Exit"]').after(dbutton);

}
