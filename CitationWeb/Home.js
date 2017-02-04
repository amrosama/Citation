(function () {
    "use strict";

    var messageBanner;
    var dataValue = "";
    var lastSelectedCitID = -1;
    var SelectedMatter;
    var rowValueGlobal;
    var dataValue_ = "";
    var jsondata;
    var formA; var formB; var formC;
    var loggedUserId;
    var mattersArr = new Array();
    var generatedID;
    var idGenerated = 0;

  
    Office.initialize = function (reason) {
  // The initialize function must be run each time a new page is loaded.
        $(document).ready(function () {
            getMainPage();
        });
    };
    function getMainPage(cb) {//this is to load Main.html page
        $.get("Main.html", function (data) {
            $("body").html(data);
            afterMainPageLoad();
            if (cb) {
                cb();
            }
        });
    }
    function getCitationPage() {//this is to load citations.html page
        $.get("citations.html", function (data) {
            $("body").html(data);
            //check if user selecting data
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    write('Action failed. Error: ' + asyncResult.error.message);
                }
                else                 
                    if (asyncResult.value.length == 0)
                        aftercitationsPageLoad();
                    else {
                        //case of user selecting text then pressed new button
                        generatedID = fillCitFormWithSelectedData(asyncResult.value);
                        //create binding
                        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: String(generatedID) }, function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                //write('Action failed. Error: ' + asyncResult.error.message);
                            } else {//success case
                                //write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
                            }
                        });
                    }
            });
        });
    }
    function citationCorrection(value) {//this is to apply ALWD rules on provided citation
        if (onlineCitation && value.indexOf("No.")==-1)
            {var v=value.split("(")[0]
            value = v.split(",")[0] + ', No.' + v.split(",")[1] +','+ v.split(",")[2]+','+v.split(",")[3]+'('+value.split("(")[1]
        }
        Office.context.document.setSelectedDataAsync(value, function (asyncResult) { });
        //generatedID
        //return value;
    }
    function IsOnlineCitation(value) {//this is to set onlineCitation variable/flag to true or false so to be used system wide
        var commas = value.split("(")[0].split(",").length
        if (commas>3) {//check for online case
            var part2parts = value.split(",")[1].trim().split(" ").length
            if (value.split(",")[1].trim().length > 3)
                onlineCitation = true;
                return;
        }
        onlineCitation=false
    }
    var onlineCitation = false;
    function getVolAbbIPN(value, value_forOnlineCit) {//this is to get Reporter Volume, Reporter Abbreviation and Initial Page Number fields from selected text/citation which is copied from legal website
        var values = value.split(" ");
        var results = new Array(3);
        if (onlineCitation) { //if this citation is online type, then no need to generate Reporter Volume , Reporter Abbreviation ,Initial Page Number
            results[0] = results[1] = results[2] = ""
            results[1] = value_forOnlineCit.trim();//this gets Database Identifier
            results[2] = value.trim();//this gets Docket Number
            return results     }
        results[0] = values[1]
        var temp = "";
        for (var i = 2; i < values.length - 1; i++) temp += values[i] + " "
        results[1] = String(temp).trim()
        results[2] = values[values.length - 1]
        return results
    }
    function getPPNcAbbDate(value, value_forOnlineCit) {//this is to get Pinpoint page number, Court Abbreviation and date from selected text/citation which is copied from legal website
        if (!value) return null;//case of value is undefind, to solve a bug
        var values = value.split(" ");
        var results = new Array(3);
        if (!onlineCitation) { results[0] = values[1] }
        else {        results[0] = value_forOnlineCit.split("(")[0].trim()        }
        values = value.split("(");
        var values2
        if (!onlineCitation)
            values2 = values[values.length - 1].split(" ");
        else
            values2 = value_forOnlineCit.split(" ");
        var cortAbb=""
        for (var i = 0; i < values2.length - 1; i++) {
            if (values2[i] == "App.") {
                cortAbb += "Ct. " + values2[i] + " ";
                if (values2[i + 2] == "Dist.")
                { values2[i + 1] = ""; values2[i + 2] = ""; }//skipping any number between App. and Dist. and skipping Dist. too
            }
            else if (values2[i] == "Super.")
                cortAbb += values2[i] + " Ct."
                //   else if (values2[i] != "1st" && values2[i] != "2d" && values2[i] != "3d" && values2[i] != "4th" && values2[i] != "5th" && values2[i] != "6th" && values2[i] != "Dist.")
                //     cortAbb += values2[i] + " ";
            else
                cortAbb += values2[i] + " ";
        }
        if (!onlineCitation)
            cortAbb = String(cortAbb).trim();
        else {
            var temp = String(cortAbb).split("(")[1].split(" ");
            cortAbb = temp[0] + ' ' + temp[1] + ' ' + temp[2];
        }
        results[1] = cortAbb
        if (!onlineCitation)
            //  results[2] =   String(values2[values2.length - 1]).substring(0, values2[values2.length - 1].length - 2)
            results[2] = String(values2[values2.length - 1]).substring(0, values2[values2.length - 1].indexOf(")"));
        else {
            var num = value_forOnlineCit.split(' ').length
            results[2] = (value_forOnlineCit.split(' ')[num - 4] + ' ' + value_forOnlineCit.split(' ')[num - 3] + ' ' + value_forOnlineCit.split(' ')[num - 1]).split(')')[0]
        }
        return results
    }
    function fillCitFormWithSelectedData(selectedValue) {//this is to create and fill citation from from selected text/citation which is coppid from legal website 
        IsOnlineCitation(selectedValue)
        var values = selectedValue.split(",");
        selectedValue = citationCorrection(selectedValue)
        var generatedID=generateCitNewID();
        //build citation edit screen by code from selected data
      
        //IsOnlineCitation(selectedValue)
        var VolAbbIPN = getVolAbbIPN(values[1], values[2]);
        var PPNcAbbDate = getPPNcAbbDate(values[2], values[3] + ', ' + values[4]);
        if (!onlineCitation)
            document.getElementById('citationList').innerHTML = '<div style="text-align:center">' + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Number</label> <input id="NumberField" class="ms-TextField-field" type="text" style=" width:93%" value="' + generatedID + '"></div>'
            + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-119px;">Case Name</label> <input id="CaseNameField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[0] + '"></div>'
            + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Reporter</label> <input id="ReporterField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[0] + '"></div>'
            + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Abrev</label> <input id="AbrevField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[1] + '"></div>'
            + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-120px;">Initial Page</label> <input id="InitialPageField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[2] + '"></div>'
             + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-125px;">Pin Point</label> <input id="PinPointField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[0] + '"></div>'
              + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Court</label> <input id="CourtField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[1] + '"></div>'
             + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-138px;">Date</label> <input id="DateField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[2] + '"></div>'
              + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-131px;">History</label> <input id="HistoryField" class="ms-TextField-field" type="text" style=" width:93%" value="' + '' + '"></div>' + '</div>';
        else
            document.getElementById('citationList').innerHTML = '<div style="text-align:center">' + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Number</label> <input id="NumberField" class="ms-TextField-field" type="text" style=" width:93%" value="' + generatedID + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-119px;">Case Name</label> <input id="CaseNameField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[0] + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Reporter</label> <input id="ReporterField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[0] + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-100px;">Database Identifier</label> <input id="AbrevField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[1] + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-107px;">Docket Number</label> <input id="InitialPageField" class="ms-TextField-field" type="text" style=" width:93%" value="' + VolAbbIPN[2] + '"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-125px;">Pin Point</label> <input id="PinPointField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[0] + '"></div>'
                      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Court</label> <input id="CourtField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[1] + '"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-138px;">Date</label> <input id="DateField" class="ms-TextField-field" type="text" style=" width:93%" value="' + PPNcAbbDate[2] + '"></div>'
                      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-131px;">History</label> <input id="HistoryField" class="ms-TextField-field" type="text" style=" width:93%" value="' + '' + '"></div>' + '</div>';

        //link events with methods goes here
        $('#save_btn_citationScreen').click(saveCitation);
        $('#formEditor-button_citationScreen').click(formEditorCitation);
        $('#done_formEditor_btn_citationScreen').click(doneFormEditorClicked);
        $('#cancel_btn_citationScreen').click(cancelCitationClicked);
        //generate id for this new created citation from selection
        idGenerated = 1;
        return generatedID;
    }
    //old copy of this function
    //function fillCitFormWithSelectedData(selectedValue) {
    //    //build citation edit screen by code from selected data
    //    var values = selectedValue.split(",");
    //    document.getElementById('citationList').innerHTML = '<div style="text-align:center">' + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Number</label> <input id="NumberField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[0]  + '"></div>'
    //    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-119px;">Case Name</label> <input id="CaseNameField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[1] + '"></div>'
    //    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Reporter</label> <input id="ReporterField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[2] + '"></div>'
    //    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Abrev</label> <input id="AbrevField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[3] + '"></div>'
    //    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-120px;">Initial Page</label> <input id="InitialPageField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[4] + '"></div>'
    //     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-125px;">Pin Point</label> <input id="PinPointField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[5] + '"></div>'
    //      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Court</label> <input id="CourtField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[6] + '"></div>'
    //     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-138px;">Date</label> <input id="DateField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[7] + '"></div>'
    //      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-131px;">History</label> <input id="HistoryField" class="ms-TextField-field" type="text" style=" width:93%" value="' + values[8] + '"></div>' + '</div>';
    //    //link events with methods goes here
    //    $('#save_btn_citationScreen').click(saveCitation);
    //    $('#formEditor-button_citationScreen').click(formEditorCitation);
    //    $('#done_formEditor_btn_citationScreen').click(doneFormEditorClicked);
    //    $('#cancel_btn_citationScreen').click(cancelCitationClicked);
    //    //generate id for this new created citation from selection
    //    idGenerated = 1;
    //    return generateCitNewID();// to be implemented 
    //}
    function aftercitationsPageLoad() {
        //build interface upon citation edit button clicked
        var content;
        if (lastSelectedCitID == -1) {
            document.getElementById('citationList').innerHTML = '<div style="text-align:center">' + '<div class="ms-TextField"><label class="ms-Label" style="position:relative;left:-128px;">Number</label> <input id="NumberField" class="ms-TextField-field" type="text" style="width:93%"></div>'
                   + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-119px;">Case Name</label> <input id="CaseNameField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                   + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Reporter</label> <input id="ReporterField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                   + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Abrev</label> <input id="AbrevField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                   + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-120px;">Initial Page</label> <input id="InitialPageField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-125px;">Pin Point</label> <input id="PinPointField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Court</label> <input id="CourtField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-138px;">Date</label> <input id="DateField" class="ms-TextField-field" type="text" style=" width:93%"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-131px;">History</label> <input id="HistoryField" class="ms-TextField-field" type="text" style=" width:93%"></div>' + '</div>';
        }
        else {
            // $.getJSON('jsonSampleV4.json', function (data) {
            for (var i = 0; i < mattersArr.length ; i++)
                if (SelectedMatter == mattersArr[i].Name) {
                    //build citation edit screen by code
                    document.getElementById('citationList').innerHTML = '<div style="text-align:center">' + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Number</label> <input id="NumberField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'Number') /*data.matters[i].citations[lastSelectedCitID].Number*/ + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-119px;">Case Name</label> <input id="CaseNameField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'CaseName') + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-128px;">Reporter</label> <input id="ReporterField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'Reporter') + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Abrev</label> <input id="AbrevField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'Abrev') + '"></div>'
                    + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-120px;">Initial Page</label> <input id="InitialPageField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'InitialPage') + '"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-125px;">Pin Point</label> <input id="PinPointField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'PinPoint') + '"></div>'
                      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-135px;">Court</label> <input id="CourtField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'Court') + '"></div>'
                     + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-138px;">Date</label> <input id="DateField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'Date') + '"></div>'
                      + '<div class="ms-TextField"><label class="ms-Label" style=" position:relative;left:-131px;">History</label> <input id="HistoryField" class="ms-TextField-field" type="text" style=" width:93%" value="' + getCitInfo(lastSelectedCitID, 'History') + '"></div>' + '</div>';
                    break;
                }
            // });
        }
        //link events with methods goes here
        $('#save_btn_citationScreen').click(saveCitation);
        $('#formEditor-button_citationScreen').click(formEditorCitation);
        $('#done_formEditor_btn_citationScreen').click(doneFormEditorClicked);
        $('#cancel_btn_citationScreen').click(cancelCitationClicked);

    }
    function cancelCitationClicked() {/*back to last/main screen*/        goToCitationsScreen(); }
    function CreateNewItem(NewCitID) {
        if (!onlineCitation)
            var newItem = { _id: NewCitID, formA: formA, formB: formB, formC: document.getElementById("NumberField").value, Number: document.getElementById("NumberField").value, CaseName: document.getElementById("CaseNameField").value, Reporter: document.getElementById("ReporterField").value, Abrev: document.getElementById("AbrevField").value, InitialPage: document.getElementById("InitialPageField").value, PinPoint: document.getElementById("PinPointField").value, Court: document.getElementById("CourtField").value, Date: document.getElementById("DateField").value, History: document.getElementById("HistoryField").value, Library: '', lastEditUserId: '', creatorId: loggedUserId, Other: '', formManEdit: '', type: '', version: '', userID: '', created: '', modified: '', highlighted: '', onlineCitation: onlineCitation };
        else
            var newItem = { _id: NewCitID, formA: formA, formB: formB, formC: document.getElementById("NumberField").value, Number: document.getElementById("NumberField").value, CaseName: document.getElementById("CaseNameField").value, dbIdentifier: document.getElementById("AbrevField").value, DocketNum: document.getElementById("InitialPageField").value, PinPoint: document.getElementById("PinPointField").value, Court: document.getElementById("CourtField").value, Date: document.getElementById("DateField").value, History: document.getElementById("HistoryField").value, Library: '', lastEditUserId: '', creatorId: loggedUserId, Other: '', formManEdit: '', type: '', version: '', userID: '', created: '', modified: '', highlighted: '', onlineCitation: onlineCitation };
        return newItem;
    }
    function doneFormEditorClicked() {
        jsondata = -1;
        //if user changed data for from A,B or C ,then duplicate citation in json file one with old data, other with new form(s) data
        if (document.getElementById('formAField').value != formA || document.getElementById('formBField').value != formB || document.getElementById('formCField').value != formC) {//user changed forms case, so duplicate the citation
            $.getJSON('jsonSampleV4.json', function (data) {
                for (var i = 0; i < data.matters.length ; i++)
                    if (SelectedMatter == data.matters[i].Name) {
                        var NewCitID = generateCitNewID();
                        LinkMatterCitation(data.matters[i]._id, NewCitID, i, data);
                        formA= document.getElementById('formAField').value;
                        formB= document.getElementById('formBField').value;
                        formC= document.getElementById('formCField').value;
                        data.citations.push(CreateNewItem(NewCitID));
                        // data.matters[i].citations.push(CreateNewItem());//add new citation node to json file
                        jsondata = data;
                        JsonObject = JSON.parse(JSON.stringify(data));
                        getMatterCitations(mattersArr[i]._id);
                        break;
                    }
                //update json file on the server side
                $.post('https://localhost:44305/api/upload/reciveJson', { json: JSON.stringify(data) }).fail(function myfunction(error) {
                    console.log('fail');
                });
            });
        }
        //back to last/main screen
        goToCitationsScreen();
    }
    function goToCitationsScreen() {
        generateCitationlistFromLocalJasonFile(rowValueGlobal, function (data) {
            //var jsondata = data;
            if (jsondata == null || jsondata == -1) jsondata = data;
            getMainPage(function () {
                for (var i = 0; i < JsonObject.matters.length ; i++)
                    if (JsonObject.matters[i].Name == rowValueGlobal)
                        for (var ii = 0; ii < JsonObject.matters[i].citations.length; ii++) {
                            document.getElementById('citationsListDiv').innerHTML += '<div style="margin-bottom:5px;">';
                            document.getElementById('citationsListDiv').innerHTML += '<p id="citItemA' + citIDs[ii] + '" class="citationList"  class="ms-font-xl" style="margin:0;">' + getCitInfo(citIDs[ii], 'formA') /*getCitInfo(citIDs[ii], 'formA')*/ + '<div id="3buttons" style="position:relative;float:right;top:-20px;"><input type="image" id="add_button' + citIDs[ii] + '" src="Images/add.png" class="_3buttons add_buttonClass" /><input type="image" id="edit_button' + citIDs[ii] + '" class="_3buttons edit_buttonClass" src="Images/edit.png" /><input type="image" id="delete_button' + citIDs[ii] + '" src="Images/del.png" class="_3buttons delete_buttonClass" /></p></div>';
                            document.getElementById('citationsListDiv').innerHTML += '<p id="citItemB' + citIDs[ii] + '"  class="ms-font-xl" style="font-size:90%;margin:0;">' + getCitInfo(citIDs[ii], 'formB') + '</p>';
                            document.getElementById('citationsListDiv').innerHTML += '<p  id="citItemC' + citIDs[ii] + '" class="ms-font-mi" style="margin:0;color:lightgray;">' + getCitInfo(citIDs[ii], 'formC') + '</p></div>';
                            document.getElementById('citationsListDiv').innerHTML += '<div id="h' + citIDs[ii] + '"> <hr/> </div>';
                        }
                //show hand cursor when user hover over citation from A text
                var citList = document.getElementsByClassName('citationList');
                for (var i = 0; i < citList.length; i++) citList[i].style.cursor = 'pointer';
                //bind events
                bindSmallButtons();
                document.getElementById('correctLegalDiv').style.display = 'none';//hide all login controls
                document.getElementById('mattersListDiv').style.display = 'none';
                document.getElementById('citationsDiv').style.display = 'inline-block';
                document.getElementById('loginDiv').style.display = 'none';//hide all login controls
            });
        });
    }
    function formEditorCitation() {
        //hide citation div
        document.getElementById('citationList').style.display = 'none';
        document.getElementById('citationButtonsDiv').style.display = 'none';
        //show form editor div
        document.getElementById('formEditorDiv').style.display = 'inline-block';
        //hide this button
        document.getElementById('formEditor-button_citationScreen').style.display = 'none';
    }
    function afterMainPageLoad() {
        // Initialize the FabricUI notification mechanism and hide it
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner = new fabric.MessageBanner(element);
        messageBanner.hideBanner();

        // If not using Word 2016, use fallback logic.
        if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
            $("#template-description").text("This sample displays the selected text.");
            $('#button-text').text("Display!");
            $('#button-desc').text("Display the selected text");
            //$('#button-insert-text').text("Insert");

            $('#highlight-button').click(displaySelectedText);
            return;
        }

        $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
        $('#button-text').text("Highlight!");
        $('#button-citation-text').text("Citation test");
        $('#button-desc').text("Highlights the longest word.");
        $('#button-sign-text').text("Sign In");


        //loadSampleData();

        //Initializing parse.com 
        Parse.initialize("ASDFGHJKL12345LEL5MDLEDTSOALDOSPDOSDELCE");
        Parse.serverURL = 'https://correctlegal.herokuapp.com/parse'

        //$('#citation-test').click(replaceTest);
        // Add a click event handler for the highlight button.
        $('#highlight-button').click(generateCitationlistFromLocalJasonFile);
        //$('#add_button').click(hightlightLongestWord)
        $('#Signin-button').click(signin_clicked);
        //$('#afterSelectionBton').click(go2citationScreen)
        $('#insert-button').click(insertCitationClicked);
        $('#mainFormat_btn').click(replaceCitationsFullDoc);
        $('#mainNew_btn').click(newButtonClicked);
        $('#mainSelect_btn').click(selectButtonClicked);
        $('#test_btn').click(test);
        $('#search_btn').click(SearchClicked)

    }
    function SearchClicked() {
            document.getElementById('citationsListDiv').innerHTML = '';//cleaning citation list div for new search result
            for (var i = 0; i < citIDs.length ; i++) {
                if (getCitInfo(citIDs[i], 'formA').indexOf(document.getElementById("search_field").value) != -1 ||
                                getCitInfo(citIDs[i], 'formB').indexOf(document.getElementById("search_field").value) != -1 ||
                                getCitInfo(citIDs[i], 'formC').indexOf(document.getElementById("search_field").value) != -1) {
                    document.getElementById('citationsListDiv').innerHTML += '<div style="margin-bottom:5px;">';
                    document.getElementById('citationsListDiv').innerHTML += '<p id="citItemA' + citIDs[i] + '" class="citationList"  class="ms-font-xl" style="margin:0;">' + getCitInfo(citIDs[i], 'formA')  + '<div id="3buttons" style="position:relative;float:right;top:-20px;"><input type="image" id="add_button' + citIDs[i] + '" src="Images/add.png" class="_3buttons add_buttonClass" /><input type="image" id="edit_button' + citIDs[i] + '" class="_3buttons edit_buttonClass" src="Images/edit.png" /><input type="image" id="delete_button' + citIDs[i] + '" src="Images/del.png" class="_3buttons delete_buttonClass" /></p></div>';
                    document.getElementById('citationsListDiv').innerHTML += '<p id="citItemB' + citIDs[i] + '"  class="ms-font-xl" style="font-size:90%;margin:0;">' + getCitInfo(citIDs[i], 'formB')  + '</p>';
                    document.getElementById('citationsListDiv').innerHTML += '<p  id="citItemC' + citIDs[i] + '" class="ms-font-mi" style="margin:0;color:lightgray;">' + getCitInfo(citIDs[i], 'formC')  + '</p></div>';
                    document.getElementById('citationsListDiv').innerHTML += '<div id="h' + citIDs[i] + '"> <hr/> </div>';
                }
            }
            //show hand cursor when user hover over citation from A text
            var citList = document.getElementsByClassName('citationList');
            for (var i = 0; i < citList.length; i++) citList[i].style.cursor = 'pointer';

            //bind events
            bindSmallButtons();
    }
    //function SearchClicked() {
    //   
    //        var CitObj = Parse.Object.extend("citations");
    //        var formAQuery = new Parse.Query(CitObj); var formBQuery = new Parse.Query(CitObj); var formCQuery = new Parse.Query(CitObj);
    //        var query = Parse.Query.or(formAQuery, formBQuery, formCQuery)
    //        //    query.startsWith("reporter", "rep");//if this line not exists, then will get all records 
    //        formAQuery.contains("formA", document.getElementById("search_field").value); formBQuery.contains("formB", document.getElementById("search_field").value); formCQuery.contains("formC", document.getElementById("search_field").value);
    //        query.find({
    //            success: function (results) {//results contins rows 
    //                generateCitationlistFromParseSearchResult(results);
    //                //for (var i = 0; i < results.length; i++) {
    //                //    var row = results[i];
    //                //    var reporter = row.get("reporter");//reporter is the column
    //                //}
    //            },
    //            error: function (error) {
    //                // alert("Error: " + error.code + " " + error.message);
    //            }
    //        });
    //}
    //function generateCitationlistFromParseSearchResult(result) {
    //    document.getElementById('citationsListDiv').innerHTML = '';//cleaning citation list div for new search result
    //    for (var i = 0; i < result.length; i++) {
    //        var row = result[i];
    //        document.getElementById('citationsListDiv').innerHTML += '<div style="margin-bottom:5px;">';
    //        document.getElementById('citationsListDiv').innerHTML += '<p id="citItemA' + row.get("objectId") /*+ citIDs[ii] */ + '" class="citationList"  class="ms-font-xl" style="margin:0;">' + row.get("formA") + '<div id="3buttons" style="position:relative;float:right;top:-20px;"><input type="image" id="add_button' + row.get("objectId") /*+ citIDs[ii]*/ + '" src="Images/add.png" class="_3buttons add_buttonClass" /><input type="image" id="edit_button' +row.get("objectId")/*citIDs[ii]*/ + '" class="_3buttons edit_buttonClass" src="Images/edit.png" /><input type="image" id="delete_button' + /*citIDs[ii]*/ row.get("objectId")+ '" src="Images/del.png" class="_3buttons delete_buttonClass" /></p></div>';
    //        document.getElementById('citationsListDiv').innerHTML += '<p id="citItemB' + row.get("objectId") /*citIDs[ii]*/ + '"  class="ms-font-xl" style="font-size:90%;margin:0;">' + row.get("formB") + '</p>';
    //        document.getElementById('citationsListDiv').innerHTML += '<p id="citItemC' + row.get("objectId") /*citIDs[ii]*/ + '" class="ms-font-mi" style="margin:0;color:lightgray;">' + row.get("formC") + '</p></div>';
    //        document.getElementById('citationsListDiv').innerHTML += '<div id="h' + row.get("objectId")/*+ citIDs[ii]*/ + '"> <hr/> </div>';
    //    }
    //    //show hand cursor when user hover over citation from A text
    //    var citList = document.getElementsByClassName('citationList');
    //    for (var i = 0; i < citList.length; i++) citList[i].style.cursor = 'pointer';

    //    //bind events
    //    bindSmallButtons();
    //}
    function test() {
        //change binding object test
        Office.context.document.bindings.getByIdAsync(String(generatedID), function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            }
            else {//success case
                //write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
                //write value
                asyncResult.value.setDataAsync('replaced value', function (Result) {
                    if (Result.status == Office.AsyncResultStatus.Failed) {
                        write('Action failed. Error: ' + asyncResult.error.message);
                    }
                });
            }
        });
        //read value
        //Office.select("bindings#" + String(generatedID)).getDataAsync(function (asyncResult) {
        //    var mybindvalue = asyncResult.value;
        //})
    }
    function newButtonClicked() {
        lastSelectedCitID = -1;
        getCitationPage();
    }
    function selectButtonClicked() {//arrow button 
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
         { valueFormat: "unformatted", filterType: "all" },
         function (asyncResult) {
             var error = asyncResult.error;
             if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 showNotification(error.name + ": " + error.message);
             }
             else {// Get selected data.
                 var selectedData = asyncResult.value;
                 dataValue_ = selectedData;
                 if (dataValue_.length == 0) {
                     //user didn't select any data case, nothing happends
                     showNotification('Please select some text, then try again ');
                 }
                 else {
                     //$.getJSON('jsonSample.json', function (data) {
                         //for (var i = 0; i < data.matters.length ; i++)
                         //    if (SelectedMatter == JsonObject.matters[i].Name)
                                 for (var ii = 0; ii < JsonObject.citations.length; ii++) {
                                     if (dataValue_.indexOf(JsonObject.citations[ii].formA) !== -1)
                                         for (var v = 0; v < $('[id^=citItem]').length; v++) {
                                             if ($('[id^=citItem]')[v].innerText == JsonObject.citations[ii].formA) {
                                                 lastSelectedCitID = $('[id^=citItem]')[v].id.replace('citItemA', '').trim();
                                                 getCitationPage();
                                                 return true;
                                             }
                                 //        }
                                 }
                     }//);
                 }
             }
         });
    }
    function replaceCitationsFullDoc() {
        getFileData(fileReady);//get all data in file, then call fileReady
    }
    var fileReady = function () {//called from replaceCitationsFullDoc
        //Your Code Go Here
        clearAllText(insert_text);
        //insert_text();
        
    }

    function saveCitation() {
        //validation for case when user try to create new citation and didn't insert any data at all
        if (lastSelectedCitID == -1 && document.getElementById('NumberField').value.length == 0 && document.getElementById('DateField').value.length == 0 && document.getElementById('HistoryField').value.length == 0 && document.getElementById('CourtField').value.length == 0 && document.getElementById('InitialPageField').value.length == 0 && document.getElementById('AbrevField').value.length == 0 && document.getElementById('ReporterField').value.length == 0 && document.getElementById('PinPointField').value.length == 0 && document.getElementById('CaseNameField').value.length == 0)
        {
            document.getElementById('errorMessage').innerHTML = 'You must insert information to create the citation, please try again';
            return;
        } else { document.getElementById('errorMessage').innerHTML = '';}
        if (lastSelectedCitID == -1) {//new citation case
            $.getJSON('jsonSampleV4.json', function (data) {
                for (var i = 0; i < data.matters.length ; i++)
                    if (SelectedMatter == data.matters[i].Name) {
                        if (!onlineCitation) {
                            formA = document.getElementById("CaseNameField").value + ', ' + document.getElementById("ReporterField").value + ' ' + document.getElementById("AbrevField").value + ' ' + document.getElementById("InitialPageField").value + ', ' + document.getElementById("PinPointField").value + ' (' + document.getElementById("CourtField").value + ' ' + document.getElementById("DateField").value + ').';// + document.getElementById("HistoryField").value + ',' + document.getElementById("HistoryField").value;
                            formB = document.getElementById("CaseNameField").value.split(" v.")[0] + ', ' + document.getElementById("ReporterField").value + ' ' + document.getElementById("AbrevField").value + ' at ' + document.getElementById("PinPointField").value + '.';
                            formC = document.getElementById("NumberField").value;
                        } else {
                            formA = document.getElementById("CaseNameField").value + ', No. ' + document.getElementById("InitialPageField").value +', '+ document.getElementById("AbrevField").value + ', ' + document.getElementById("PinPointField").value + ' (' + document.getElementById("CourtField").value + ' ' + document.getElementById("DateField").value + ').'
                            formB = document.getElementById("CaseNameField").value.split(" v.")[0] + ', ' + document.getElementById("AbrevField").value + ', ' + document.getElementById("PinPointField").value + '.';
                            formC = document.getElementById("NumberField").value;
                            
                        }
                        var NewCitID;
                        if (idGenerated == 0)
                            NewCitID = generateCitNewID();
                        else {
                            NewCitID = generatedID
                            idGenerated=0
                        }
                        LinkMatterCitation(data.matters[i]._id, NewCitID, i,data);
                        data.citations.push(CreateNewItem(NewCitID));
                        //set formA,B,C in fromEditor div
                        document.getElementById('formAField').value = formA;
                        document.getElementById('formBField').value = formB;
                        document.getElementById('formCField').value = formC;
                        JsonObject = JSON.parse(JSON.stringify(data));
                        getMatterCitations(data.matters[i]._id);
                        break;
                    }
                //update json file on the server side
                $.post('https://localhost:44305/api/upload/reciveJson', { json: JSON.stringify(data) }).fail(function myfunction(error) {
                    console.log('fail');
                });
            });
        }
        else {//editing exists citation case
            $.getJSON('jsonSampleV4.json', function (data) {
                for (var ii = 0; ii < data.citations.length; ii++) {
                    if (data.citations[ii]._id == lastSelectedCitID) {
                        //return eval('JsonObject.citations[i].' + info);
                        formA = getCitInfo(lastSelectedCitID, 'formA');// data.matters[i].citations[lastSelectedCitID].formA;
                        formB = getCitInfo(lastSelectedCitID, 'formB');//data.matters[i].citations[lastSelectedCitID].formB;
                        formC = getCitInfo(lastSelectedCitID, 'formC');// data.matters[i].citations[lastSelectedCitID].formC;
                        data.citations[ii].Number = document.getElementById("NumberField").value;
                        data.citations[ii].CaseName = document.getElementById("CaseNameField").value;
                        //data.matters[i].citations[lastSelectedCitID].Reporter = document.getElementById("ReporterField").value;
                        data.citations[ii].Reporter = document.getElementById("ReporterField").value;
                        data.citations[ii].Abrev = document.getElementById("AbrevField").value;
                        data.citations[ii].InitialPage = document.getElementById("InitialPageField").value;
                        data.citations[ii].PinPoint = document.getElementById("PinPointField").value;
                        data.citations[ii].Court = document.getElementById("CourtField").value;
                        data.citations[ii].Date = document.getElementById("DateField").value;
                        data.citations[ii].History = document.getElementById("HistoryField").value;
                        document.getElementById('Message').innerHTML = 'Saved Successfuly';
                        //set formA,B,C in fromEditor div
                        document.getElementById('formAField').value = formA;
                        document.getElementById('formBField').value = formB;
                        document.getElementById('formCField').value = formC;
                        JsonObject = JSON.parse(JSON.stringify(data));
                        for (var i = 0; i < data.matters.length ; i++)
                            if (SelectedMatter == data.matters[i].Name)
                                getMatterCitations(data.matters[i]._id);
                        break;
                    }
                }
                //update json file on the server side
                $.post('https://localhost:44305/api/upload/reciveJson', { json: JSON.stringify(data) }).fail(function myfunction(error) {
                    console.log('fail');
                });
            });
        }

        //show form editor button
        document.getElementById('formEditor-button_citationScreen').style.display = 'inline-block';
    }
    function generateCitNewID() {//this function need to be re-implemented upon the logic we need to generate new id for new created citation
        return Math.floor((Math.random() * 1000000000) + 1);//generate random number between 1 and 1000000000
    }
    function LinkMatterCitation(matterID,citationID,matterIndex,data) {//this function taks matter id and citation id (new created citation id) to link it together in json file
        //$.getJSON('jsonSampleV4.json', function (data) {
            var alreadyExists=0;
            for (var i = 0; i < JsonObject.userMatters.length ; i++) 
                if (JsonObject.userMatters[i].userID == loggedUserId && JsonObject.userMatters[i].matterId) alreadyExists = 1;
            if (!alreadyExists)  data.userMatters.push({ userID: loggedUserId, matterId: data.matters[matterIndex]._id });
            data.matters[matterIndex].citations.push({ cid: citationID });
                //update json file on the server side
                //$.post('https://localhost:44305/api/upload/reciveJson', { json: JSON.stringify(data) }).fail(function myfunction(error) {
                //    console.log('fail');
                //});
        //});
    }
    function insertCitationClicked() {
        if (lastSelectedCitID == -1) {//don't continue, user didn't select any citation
            showNotification('Please select citation first');
        } else {
            //$.getJSON('jsonSample.json', function (data) {
            for (var i = 0; i < JsonObject.matters.length ; i++) {
                if (SelectedMatter == JsonObject.matters[i].Name ) {
                    Office.context.document.setSelectedDataAsync(getCitInfo(lastSelectedCitID, 'formB'),
                       { coercionType: Office.CoercionType.Text });
                        break;
                    }
                }
            //});
            showNotification('');
        }
    }
    function go2citationScreen() {
        window.location.href = 'citations.html';
    }
    function generateListofMatters() {
            var content = '<table id="mattersTbl" class="display" cellspacing="0"><thead><tr><th>Matter</th></tr></thead><tfoot><tr><th>Select Matter to show related citations</th> </tr> </tfoot><tbody>';
            for (var i = 0; i < mattersArr.length ; i++) { content += '<tr> <td>' + mattersArr[i].Name + '</td> </tr>'; }
            content += ' </tbody> </table>';
            document.getElementById('mattersListDiv').innerHTML = content;
            //bind events
            initializeMatterTable();
    }
    //old code before jason V4
    //function generateListofMatters() {
    //    $.getJSON('jsonSample.json', function (data) {
    //        var content = '<table id="mattersTbl" class="display" cellspacing="0"><thead><tr><th>Matter</th></tr></thead><tfoot><tr><th>Select Matter to show related citations</th> </tr> </tfoot><tbody>';
    //        for (var i = 0; i < data.matters.length ; i++) { content += '<tr> <td>' + data.matters[i].Name + '</td> </tr>'; }
    //        content += ' </tbody> </table>';
    //        document.getElementById('mattersListDiv').innerHTML = content;
    //        //bind events
    //        initializeMatterTable();
    //    });
    //}
    function initializeMatterTable() {
        var table = $('#mattersTbl').DataTable();
        $('#mattersTbl tbody').on('click', 'tr', function () {
            $('#mattersTbl tbody tr.selected').removeClass('selected');
            $(this).toggleClass('selected');
            var rowdata = table.row('.selected').data();
            SelectedMatter = rowdata[0];
            rowClicked(rowdata[0]);
        });

    }
    function rowClicked(rowValue) {
        generateCitationlistFromLocalJasonFile(rowValue);
        document.getElementById('correctLegalDiv').style.display = 'none';//hide all login controls
        document.getElementById('mattersListDiv').style.display = 'none';
        document.getElementById('citationsDiv').style.display = 'inline-block';
    }
    //function getMatterCitLength(matterId) {
    //    for (var i = 0; i < JsonObject.matters.length; i++)
    //        if (JsonObject.matters[i]._id == matterId)
    //            return JsonObject.matters[i].citations.length;
    //}
    var citIDs = Array();//where we save citation ids for matter passed in the parameter matterId
    function getMatterCitations(matterId) {
        citIDs = new Array();
        var index=0;
        for (var i = 0; i < JsonObject.matters.length; i++)
            if (JsonObject.matters[i]._id == matterId) {
                for (var ii = 0; ii < JsonObject.matters[i].citations.length ; ii++)
                    citIDs[index++] = JsonObject.matters[i].citations[ii].cid;
                break;
            }
    }
    function getCitInfo(citId,info)
    {
        for (var i = 0; i < JsonObject.citations.length; i++){
            if (JsonObject.citations[i]._id == citId)
                return eval('JsonObject.citations[i].' + info);
        }
        return 'Requested data not exists in Json file';//after changing backend to parse instead of json we will get this message
    }
    function generateCitationlistFromLocalJasonFile(rowValue, cb) {
        rowValueGlobal = rowValue;
        //$.getJSON('jsonSample.json', function (data) {
            if (cb) {
                // cb(data);
                cb();
            } else {
                for (var i = 0; i < mattersArr.length ; i++)
                    if (mattersArr[i].Name == rowValue){  getMatterCitations(mattersArr[i]._id);
                        for (var ii = 0; ii < /*getMatterCitLength(mattersArr[i]._id)*/citIDs.length ; ii++) {
                            document.getElementById('citationsListDiv').innerHTML += '<div style="margin-bottom:5px;">';
                            document.getElementById('citationsListDiv').innerHTML += '<p id="citItemA' + citIDs[ii] + '" class="citationList"  class="ms-font-xl" style="margin:0;">' +getCitInfo(citIDs[ii],'formA') /*data.matters[i].citations[ii].formA*/ + '<div id="3buttons" style="position:relative;float:right;top:-20px;"><input type="image" id="add_button' + citIDs[ii] + '" src="Images/add.png" class="_3buttons add_buttonClass" /><input type="image" id="edit_button' + citIDs[ii] + '" class="_3buttons edit_buttonClass" src="Images/edit.png" /><input type="image" id="delete_button' + citIDs[ii] + '" src="Images/del.png" class="_3buttons delete_buttonClass" /></p></div>';
                            document.getElementById('citationsListDiv').innerHTML += '<p id="citItemB' + citIDs[ii] + '"  class="ms-font-xl" style="font-size:90%;margin:0;">' +getCitInfo(citIDs[ii],'formB') /*data.matters[i].citations[ii].formB*/ + '</p>';
                            document.getElementById('citationsListDiv').innerHTML += '<p  id="citItemC' + citIDs[ii] + '" class="ms-font-mi" style="margin:0;color:lightgray;">' +getCitInfo(citIDs[ii],'formC') /*data.matters[i].citations[ii].formC*/ + '</p></div>';
                            document.getElementById('citationsListDiv').innerHTML += '<div id="h' + citIDs[ii] + '"> <hr/> </div>';
                        }}
                //show hand cursor when user hover over citation from A text
                var citList = document.getElementsByClassName('citationList');
                for (var i = 0; i < citList.length; i++) citList[i].style.cursor = 'pointer';

                //bind events
                bindSmallButtons();
            }
        //});
    }
    function bindSmallButtons() {
        $('.add_buttonClass').click(add_buttons_clicked); //hightlightLongestWord);
        $('.edit_buttonClass').click(edit_buttons_clicked);
        $('.delete_buttonClass').click(delete_buttons_clicked);
        $('.citationList').click(divClicked);
    }
    function divClicked() {
        //document.getElementById('testDiv').innerHTML += this.id+' clicked ';
        //unselect/make all font color black
        $('.citationList').css('color', 'black');
        document.getElementById(this.id).style.color = "blue";
        lastSelectedCitID = (this.id.replace('citItemA', '')).trim();
    }
    function add_buttons_clicked() {
        //document.getElementById('testDiv').innerHTML += 'add button clicked';
        lastSelectedCitID = this.id.replace('add_button', '');

    }
    function edit_buttons_clicked() {
        lastSelectedCitID = (this.id.replace('edit_button', '')).trim();
        $.get("citations.html", function (data) {
            $("body").html(data);
            aftercitationsPageLoad();
        });
    }
    function delete_buttons_clicked() {
        //document.getElementById('testDiv').innerHTML += 'delete button clicked';
        lastSelectedCitID = (this.id.replace('delete_button', '')).trim();
        $.getJSON('jsonSampleV4.json', function (data) {
            //for (var i = 0; i < data.matters.length ; i++)
            //    if (SelectedMatter == data.matters[i].Name) {
            //data.matters[i].citations.splice(lastSelectedCitID, 1);//remove deleted citation from json object
            for (var i = 0; i < data.citations.length; i++)
                if (data.citations[i]._id == lastSelectedCitID)
                    data.citations.splice(i, 1);//remove deleted citation from json object
              //delete cid from matters
            for (var i = 0; i < data.matters.length ; i++)
                if (SelectedMatter == data.matters[i].Name)
                    for (var v = 0; v < data.matters[i].citations.length; v++)
                        if (data.matters[i].citations[v].cid == lastSelectedCitID)
                            data.matters[i].citations.splice(v, 1);

                    //remove deleted citation from UI
                    document.getElementById('citItemA' + lastSelectedCitID).remove();
                    document.getElementById('citItemB' + lastSelectedCitID).remove();
                    document.getElementById('citItemC' + lastSelectedCitID).remove();
                    document.getElementById('delete_button' + lastSelectedCitID).remove();
                    document.getElementById('edit_button' + lastSelectedCitID).remove();
                    document.getElementById('add_button' + lastSelectedCitID).remove();
                    document.getElementById('h' + lastSelectedCitID).remove();
                    //break;
                //}
            //update json file on the server side
            //remove deleted citation from json file
            $.post('https://localhost:44305/api/upload/reciveJson', { json: JSON.stringify(data) }).fail(function myfunction(error) {
                console.log('fail');
            });
        });

    }
    function replaceTest() {//just a test, not used
        getFileData();//get all data in file
        // Office.select()
        //getText();//get selected data 
        window.setTimeout(insert_text, 5000); //lowering 5000 may trauncate data in case of a lot of text in the document 
        clearAllText();
        insert_text();
    }
    function login() {
        //login
        var User = Parse.Object.extend("User");
   //   if ($('#usernameField').val() == 'a' && $('#passwordField').val() == '1') {
        Parse.User.logIn(document.getElementById('usernameField').value, document.getElementById("passwordField").value, {
            success: function (result) {
                    document.getElementById('loginDiv').style.display = 'none';//hide all login controls
                    document.getElementById('correctLegalDiv').style.display = 'inline-block';
                    generateListofMatters();
                //var username = result.get("username");
                //var createDate = result.createdAt;
                //var updateDate = result.updatedAt;
                //var id = result.id;
            },
            error: function (error) {
                ShowSignInButton();
                document.getElementById('loginMessageDIV').innerHTML = 'Sorry, the username and password you entered do not match. Please try again.';
                //alert("Error: " + error.code + " " + error.message);
            }
        });    
    }
    function signin_clicked() {
        document.getElementById('Signin-button').style.display = 'none';
        document.getElementById('signing1').style.display = 'inline-block'; //or //link.style.visibility = 'hidden';
        document.getElementById('signing2').style.display = 'inline-block';
        setTimeout(login, 700);
        //here after calling api to validate user and allow him to log-in, we will get his id
        //for now we will set static id for the user
        loggedUserId = "aaa";

        $.getJSON('jsonSampleV4.json', function (data) {
            JsonObject = JSON.parse(JSON.stringify(data));
            getUserMatters();
        });
       
    }
    var JsonObject;
    function getUserMatters() {
        //if (JsonObject == null) { getUserMatters(useId);        }
        var matterIdsArr = new Array();
        var index = 0; var index2 = 0;
        //$.getJSON('jsonSampleV4.json', function (data) {
        //var test = JsonObject.matters.length;
        for (var i = 0; i < JsonObject.userMatters.length; i++)
            if (JsonObject.userMatters[i].userID == loggedUserId) matterIdsArr[index++] = JsonObject.userMatters[i].matterId;
            
            for (var i = 0; i < matterIdsArr.length; i++)
                for (var u = 0; u < JsonObject.matters.length; u++)
                    if (matterIdsArr[i] == JsonObject.matters[u]._id)
                        mattersArr[index2++] = { "_id": JsonObject.matters[u]._id, "Name": JsonObject.matters[u].Name };
    }
    function ShowSignInButton() {
        document.getElementById('signing1').style.display = 'none';
        document.getElementById('signing2').style.display = 'none';
        document.getElementById('Signin-button').style.display = 'inline-block';
    }
    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }
    function clearAllText(cb) {
            // Run a batch operation against the Word object model.
            Word.run(function (context) {
                // Create a proxy object for the document body.
                var body = context.document.body;
                // Queue a commmand to clear the contents of the body.
                body.clear();
                // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
                return context.sync();
            })
                .then(function () {
                    cb();
                })
            .catch(errorHandler);
        
    }
    function getText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "unformatted", filterType: "all" },
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    write(error.name + ": " + error.message);
                }
                else {
                    // Get selected data.
                    dataValue = asyncResult.value;
                    //write('Selected data is ' + dataValue);
                    //   showNotification(dataValue);
                }
            });
    }
    //function to replace all instead of just replace one
    String.prototype.replaceAll = function (find, replace) {
        var str = this;
        return str.replace(new RegExp(find, 'g'), replace);
    };
    function insert_text() {
        var citLocs = new Array();
        var cutCharCount;
        var x=0;
        for (var i = 0; i < JsonObject.matters.length ; i++)
            if (JsonObject.matters[i].Name == rowValueGlobal) {
                for (var ii = 0; ii < JsonObject.matters[i].citations.length; ii++) {
                    var tempData = dataValue;
                    cutCharCount = 0;
                    //get data of each occurance for each citation                   
                    var locAlength = getCitInfo(citIDs[ii], 'formA').length;
                    var locBlength = getCitInfo(citIDs[ii], 'formB').length;
                    var locClength = getCitInfo(citIDs[ii], 'formC').length;
                    do {
                        var locA = tempData.indexOf(getCitInfo(citIDs[ii], 'formA'));
                        var locB = tempData.indexOf(getCitInfo(citIDs[ii], 'formB'));
                        var locC = tempData.indexOf(getCitInfo(citIDs[ii], 'formC'));
                        if (tempData.indexOf(getCitInfo(citIDs[ii], 'formA')) == -1) locA = 9999999999;
                        if (tempData.indexOf(getCitInfo(citIDs[ii], 'formB')) == -1) locB = 9999999999;
                        if (tempData.indexOf(getCitInfo(citIDs[ii], 'formC')) == -1) locC = 9999999999;
                        if (locA + locB + locC == 9999999999 * 3)//this citation is used in this document
                            continue;
                        else {
                            citLocs.push(new Array())
                            citLocs[x].push(citIDs[ii]);                  
                        }
                        if (locA != 9999999999 && locA < locB && locA < locC) {
                            citLocs[x].push('A'); citLocs[x].push(locA + cutCharCount);
                            tempData = tempData.substr(locA + locAlength);//get rest of the string
                            cutCharCount += locA + locAlength;
                        }
                        else if (locB != 9999999999 && locB < locA && locB < locC) {
                            citLocs[x].push('B'); citLocs[x].push(locB + cutCharCount);
                            tempData = tempData.substr(locB + locBlength);//get rest of the string
                            cutCharCount += locB + locBlength;
                        }
                        else if (locC != 9999999999 && locC < locA && locC < locB) {
                            citLocs[x].push('C'); citLocs[x].push(locC + cutCharCount);
                            tempData = tempData.substr(locC + locClength);//get rest of the string
                            cutCharCount += locC + locClength;
                        }
                        x++;
                    } while (!(locA == 9999999999 && locB == 9999999999 && locC == 9999999999))
                }
                //sort citLocs array assending upon citation locations
                var citLocsSorted = new Array();
                var j=0;
                var min;
                var minIndex;
                do {
                    min = 999999999999;
                    for (var i = 0; i < citLocs.length; i++) {                   
                        if (citLocs[i][2] < min)
                        { min = citLocs[i][2]; minIndex = i; }
                    }
                    citLocsSorted.push(new Array())
                    citLocsSorted[j].push(citLocs[minIndex][0]);
                    citLocsSorted[j].push(citLocs[minIndex][1]);
                    citLocsSorted[j].push(citLocs[minIndex][2]);
                    j++;
                    citLocs.splice(minIndex, 1);
                }while (citLocs.length>0)
                //implement replacement logic to put citations in correct forms A,B & C and pinpoint too
                var takeOrgive = 0; var len; var citForm;
                for (var u = 0; u < citLocsSorted.length; u++) {
                    if (citLocsSorted[u][1] == 'A') { len = getCitInfo(citLocsSorted[u][0], 'formA').length; citForm = getCitInfo(citLocsSorted[u][0], 'formA'); }
                    else if (citLocsSorted[u][1] == 'B') { len = getCitInfo(citLocsSorted[u][0], 'formB').length; citForm = getCitInfo(citLocsSorted[u][0], 'formB'); }
                    else { len = getCitInfo(citLocsSorted[u][0], 'formC').length; citForm = getCitInfo(citLocsSorted[u][0], 'formC'); }
                    dataValue = dataValue.replace(dataValue.substring(citLocsSorted[u][2] - takeOrgive , citLocsSorted[u][2] - takeOrgive + len), getCitInfo(citLocsSorted[u][0], 'formC'))
                    takeOrgive += getCitInfo(citLocsSorted[u][0], 'formB').length - getCitInfo(citLocsSorted[u][0], 'formC').length;
                }
                
                Office.context.document.setSelectedDataAsync(dataValue,
                                     { coercionType: Office.CoercionType.Text },
                                     function (result) {  /* Access the results, if necessary.*/ });
                break;
            }    
        citLocs.length = 0;
    }
    var myFile_;//this used to be able to call closeAsync, other wise error will come up after 3rd time f using 
    // Get all of the content from a Word document in 1KB chunks of text.
    function getFileData(cb) {
        dataValue = '';//reset as this function will fill this variable 
        Office.context.document.getFileAsync(
        Office.FileType.Text,
    {
        sliceSize: 1000
    },
function (asyncResult) {
    if (asyncResult.status === 'succeeded') {
        var myFile = asyncResult.value,
          state = {
              file: myFile,
              counter: 0,
              sliceCount: myFile.sliceCount
          };
        getSliceData(state, cb);
        myFile_ = myFile;
    }
});
        if (myFile_) myFile_.closeAsync();//amr
    }
    // Get a slice from the file, as specified by
    // the counter contained in the state parameter.
    function getSliceData(state, cb) {
        state.file.getSliceAsync(
          state.counter,
      function (result) {
          var slice = result.value,
            data = slice.data;
          dataValue += slice.data;
          state.counter++;
          // Do something with the data.
          // Check to see if the final slice in the file has
          // been reached—if not, get the next slice;
          // if so, close the file.
          if (state.counter < state.sliceCount) {
              getSliceData(state, cb);
          }
          else if (state.counter === state.sliceCount) {
              cb();
          }
          
      });
        
    }

    function loadSampleData() {

        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText("This is a sample text inserted in M86 the document  A Missouri man ambushed and killed three law officers and wounded three others in Baton Rouge on Sunday during a time when police nationwide M86 and in the Louisiana city in particular have been on high alert after five officers were killed in a Dallas ambush July 7. M86 Louisiana State Police announced last week that they had received threats of plots against Baton Rouge policeOn Sunday,  M86 a man identified as Gavin Long of Kansas City went on a shooting rampage on his 29th birthday that left two police officers  M86 and a sheriff's deputy dead, police sources said. Long, who  M86 was African-American, was  M86 a former Marine who spent time in Iraq and was discharged at the rank of sergeant in 2010, according  M86 to the U.S. military.Police officers who responded to Sunday's shootings killed Long in a gunbattle after the other officers  M86 were ambushed, police  M86 sources told CNN. The murder weapon was an AR-15 style semi-automatic rifle, law enforcement sources told CNN.Police have not officially released the names of the victims but one M86  was identified  M86 by family members as Officer Montrell Jackson. Law officers Matthew Gerald and Brad Garafola were also M86  killed, according to sources close to the department. That was corroborated  M86 with social media posts.The three law-enforcement  M86 officers killed in Baton Rough, Louisiana, were, from left,  M86 Montrell Jackson,  Brad Garafola and Matthew Gerald.The  M86 three law-enforcement officers killed in Baton Rough, Louisiana, were, from left, Montrell Jackson, Brad Garafola and Matthew Gerald.The gunman also critically wounded M86  a deputy who is fighting for his life, said East  M86 Baton Rouge Parish Sheriff Sid Gautreaux. Another wounded deputy and police officer have non-life-threatening wounds, law  M86 officers said.Jackson had posted on Facebook on July 8 how  M86 physically and emotionally drained he had been since protests had erupted in Baton Rouge after the July 5 killing of  M86 Alton Sterling by police.I swear to God I love this M86  city, but I wonder if this city loves me. In uniform I  M86 get nasty, hateful looks and out of uniform some consider  M86 me a threat. ... These are trying times. Please don't let  M86 hate infect your heart.Gunman made frequent web postsLong M86  was a prolific user of social media, with dozens of  M86 videos, podcasts, tweets and posts under his pseudonym  M86 Cosmo Setepenra. Under M86  that name,  M86 Long also tweeted a link to a news story about Dallas shooter Micah Johnson and said the shooter was one of us! # MY Religion  M86 is Justice.A law enforcement source said Long was  M86 not alone during his stay in Baton Rouge, but its M86  unclear if others he was with knew about or were actively M86  involved in any plot.Police gave the name of the M86  man who shot 6 police officers in Baton Rouge on M86  July 17 as Gavin Long. Online, he used the name  M86 Cosmo  M86 Setepenra,&quot; and posted on a YouTube channel of  M86 that name.Police gave the name of the man who shot 6 police officers in Baton Rouge on July 17 as Gavin Long. Online, he used the name Cosmo Setepenra, and posted on a YouTube channel of that name.The FBI is running down names of possible associates, another law enforcement official said.In YouTube videos posted July 8 and 10, reviewed by CNN, Long, using the name Cosmo,  M86 spoke about the need for fighting back and what people should  M86 say about him if anything happened to me.In the July 10  M86 video, recorded, he said, in Dallas, he says, Zero have  M86 been successful just over simple protesting You gotta fight  M86 back, he says on the video.Two  M86 law enforcement sources tell CNN that Long rented a car  M86 in Kansas City after the Dallas shootings and drove it to Baton Rouge. Given that Long posted a YouTube video from Dallas on July 10, it is likely he drove to Baton Rouge via Dallas.Calls for end to violenceQuinyetta McMillon, mother of Sterlings son Cameron, put out a statement through her lawyers M86  condemning the ambush.We are disgusted by the despicable  M86 act of violence today that resulted in the shooting deaths M86  of members of the Baton Rouge law enforcement. My family is  M86 heartbroken for the officers and their families. ...  M86 We reject violence of any kind directed at members of law  M86 enforcement or citizens.My hope is that one day soon we can come M86  together and find solutions to the very important issues  M86 facing our nation rather than continuing to hurt one M86  another.President Barack Obama on Sunday condemned the  M86 killings and all attacks on law enforcement.We as a  M86 nation have to be loud and clear that nothing justifies violence  M86 against law enforcement, Obama said, speaking from the  M86 White House press briefing room. Attacks on police are an M86  attack on all of us and the rule of law that makes society  M86 possible. In a written statement earlier in the day, Obama M86  called the Baton Rough shootings a cowardly and reprehensible  M86 assault.'No talking, just shooting'The shooting Sunday M86  took place around 8:40 a.m. (9:40 a.m. ET) in the city of M86  about 230,000 people, already tense after the high-profile M86  police shooting of Sterling, an African-American man, on July  M86 5.On Sunday, police received a call of a suspicious person  M86 walking down Airline Highway with an assault rifle, a source  M86 with knowledge of the investigation told CNN.When  M86 police arrived, the shooting began.There was no talking,  M86 just shooting, Baton Rouge Police Cpl. L.J. McKneely said.At  M86 an afternoon news conference, local and state authorities, M86  including Louisiana Gov. John ...........  A Missouri man ambushed and killed three law officers and wounded three others in Baton Rouge on Sunday during a time when police nationwide M86 and in the Louisiana city in particular have been on high alert after five officers were killed in a Dallas ambush July 7. M86 Louisiana State Police announced last week that they had received threats of plots against Baton Rouge policeOn Sunday,  M86 a man identified as Gavin Long of Kansas City went on a shooting rampage on his 29th birthday that left two police officers  M86 and a sheriff's deputy dead, police sources said. Long, who  M86 was African-American, was  M86 a former Marine who spent time in Iraq and was discharged at the rank of sergeant in 2010, according  M86 to the U.S. military.Police officers who responded to Sunday's shootings killed Long in a gunbattle after the other officers  M86 were ambushed, police  M86 sources told CNN. The murder weapon was an AR-15 style semi-automatic rifle, law enforcement sources told CNN.Police have not officially released the names of the victims but one M86  was identified  M86 by family members as Officer Montrell Jackson. Law officers Matthew Gerald and Brad Garafola were also M86  killed, according to sources close to the department. That was corroborated  M86 with social media posts.The three law-enforcement  M86 officers killed in Baton Rough, Louisiana, were, from left,  M86 Montrell Jackson,  Brad Garafola and Matthew Gerald.The  M86 three law-enforcement officers killed in Baton Rough, Louisiana, were, from left, Montrell Jackson, Brad Garafola and Matthew Gerald.The gunman also critically wounded M86  a deputy who is fighting for his life, said East  M86 Baton Rouge Parish Sheriff Sid Gautreaux. Another wounded deputy and police officer have non-life-threatening wounds, law  M86 officers said.Jackson had posted on Facebook on July 8 how  M86 physically and emotionally drained he had been since protests had erupted in Baton Rouge after the July 5 killing of  M86 Alton Sterling by police.I swear to God I love this M86  city, but I wonder if this city loves me. In uniform I  M86 get nasty, hateful looks and out of uniform some consider  M86 me a threat. ... These are trying times. Please don't let  M86 hate infect your heart.Gunman made frequent web postsLong M86  was a prolific user of social media, with dozens of  M86 videos, podcasts, tweets and posts under his pseudonym  M86 Cosmo Setepenra. Under M86  that name,  M86 Long also tweeted a link to a news story about Dallas shooter Micah Johnson and said the shooter was one of us! # MY Religion  M86 is Justice.A law enforcement source said Long was  M86 not alone during his stay in Baton Rouge, but its M86  unclear if others he was with knew about or were actively M86  involved in any plot.Police gave the name of the M86  man who shot 6 police officers in Baton Rouge on M86  July 17 as Gavin Long. Online, he used the name  M86 Cosmo  M86 Setepenra,&quot; and posted on a YouTube channel of  M86 that name.Police gave the name of the man who shot 6 police officers in Baton Rouge on July 17 as Gavin Long. Online, he used the name Cosmo Setepenra, and posted on a YouTube channel of that name.The FBI is running down names of possible associates, another law enforcement official said.In YouTube videos posted July 8 and 10, reviewed by CNN, Long, using the name Cosmo,  M86 spoke about the need for fighting back and what people should  M86 say about him if anything happened to me.In the July 10  M86 video, recorded, he said, in Dallas, he says, Zero have  M86 been successful just over simple protesting You gotta fight  M86 back, he says on the video.Two  M86 law enforcement sources tell CNN that Long rented a car  M86 in Kansas City after the Dallas shootings and drove it to Baton Rouge. Given that Long posted a YouTube video from Dallas on July 10, it is likely he drove to Baton Rouge via Dallas.Calls for end to violenceQuinyetta McMillon, mother of Sterlings son Cameron, put out a statement through her lawyers M86  condemning the ambush.We are disgusted by the despicable  M86 act of violence today that resulted in the shooting deaths M86  of members of the Baton Rouge law enforcement. My family is  M86 heartbroken for the officers and their families. ...  M86 We reject violence of any kind directed at members of law  M86 enforcement or citizens.My hope is that one day soon we can come M86  together and find solutions to the very important issues  M86 facing our nation rather than continuing to hurt one M86  another.President Barack Obama on Sunday condemned the  M86 killings and all attacks on law enforcement.We as a  M86 nation have to be loud and clear that nothing justifies violence  M86 against law enforcement, Obama said, speaking from the  M86 White House press briefing room. Attacks on police are an M86  attack on all of us and the rule of law that makes society  M86 possible. In a written statement earlier in the day, Obama M86  called the Baton Rough shootings a cowardly and reprehensible  M86 assault.'No talking, just shooting'The shooting Sunday M86  took place around 8:40 a.m. (9:40 a.m. ET) in the city of M86  about 230,000 people, already tense after the high-profile M86  police shooting of Sterling, an African-American man, on July  M86 5.On Sunday, police received a call of a suspicious person  M86 walking down Airline Highway with an assault rifle, a source  M86 with knowledge of the investigation told CNN.When  M86 police arrived, the shooting began.There was no talking,  M86 just shooting, Baton Rouge Police Cpl. L.J. McKneely said.At  M86 an afternoon news conference, local and state authorities, M86  including Louisiana Gov. John ..............  A Missouri man ambushed and killed three law officers and wounded three others in Baton Rouge on Sunday during a time when police nationwide M86 and in the Louisiana city in particular have been on high alert after five officers were killed in a Dallas ambush July 7. M86 Louisiana State Police announced last week that they had received threats of plots against Baton Rouge policeOn Sunday,  M86 a man identified as Gavin Long of Kansas City went on a shooting rampage on his 29th birthday that left two police officers  M86 and a sheriff's deputy dead, police sources said. Long, who  M86 was African-American, was  M86 a former Marine who spent time in Iraq and was discharged at the rank of sergeant in 2010, according  M86 to the U.S. military.Police officers who responded to Sunday's shootings killed Long in a gunbattle after the other officers  M86 were ambushed, police  M86 sources told CNN. The murder weapon was an AR-15 style semi-automatic rifle, law enforcement sources told CNN.Police have not officially released the names of the victims but one M86  was identified  M86 by family members as Officer Montrell Jackson. Law officers Matthew Gerald and Brad Garafola were also M86  killed, according to sources close to the department. That was corroborated  M86 with social media posts.The three law-enforcement  M86 officers killed in Baton Rough, Louisiana, were, from left,  M86 Montrell Jackson,  Brad Garafola and Matthew Gerald.The  M86 three law-enforcement officers killed in Baton Rough, Louisiana, were, from left, Montrell Jackson, Brad Garafola and Matthew Gerald.The gunman also critically wounded M86  a deputy who is fighting for his life, said East  M86 Baton Rouge Parish Sheriff Sid Gautreaux. Another wounded deputy and police officer have non-life-threatening wounds, law  M86 officers said.Jackson had posted on Facebook on July 8 how  M86 physically and emotionally drained he had been since protests had erupted in Baton Rouge after the July 5 killing of  M86 Alton Sterling by police.I swear to God I love this M86  city, but I wonder if this city loves me. In uniform I  M86 get nasty, hateful looks and out of uniform some consider  M86 me a threat. ... These are trying times. Please don't let  M86 hate infect your heart.Gunman made frequent web postsLong M86  was a prolific user of social media, with dozens of  M86 videos, podcasts, tweets and posts under his pseudonym  M86 Cosmo Setepenra. Under M86  that name,  M86 Long also tweeted a link to a news story about Dallas shooter Micah Johnson and said the shooter was one of us! # MY Religion  M86 is Justice.A law enforcement source said Long was  M86 not alone during his stay in Baton Rouge, but its M86  unclear if others he was with knew about or were actively M86  involved in any plot.Police gave the name of the M86  man who shot 6 police officers in Baton Rouge on M86  July 17 as Gavin Long. Online, he used the name  M86 Cosmo  M86 Setepenra,&quot; and posted on a YouTube channel of  M86 that name.Police gave the name of the man who shot 6 police officers in Baton Rouge on July 17 as Gavin Long. Online, he used the name Cosmo Setepenra, and posted on a YouTube channel of that name.The FBI is running down names of possible associates, another law enforcement official said.In YouTube videos posted July 8 and 10, reviewed by CNN, Long, using the name Cosmo,  M86 spoke about the need for fighting back and what people should  M86 say about him if anything happened to me.In the July 10  M86 video, recorded, he said, in Dallas, he says, Zero have  M86 been successful just over simple protesting You gotta fight  M86 back, he says on the video.Two  M86 law enforcement sources tell CNN that Long rented a car  M86 in Kansas City after the Dallas shootings and drove it to Baton Rouge. Given that Long posted a YouTube video from Dallas on July 10, it is likely he drove to Baton Rouge via Dallas.Calls for end to violenceQuinyetta McMillon, mother of Sterlings son Cameron, put out a statement through her lawyers M86  condemning the ambush.We are disgusted by the despicable  M86 act of violence today that resulted in the shooting deaths M86  of members of the Baton Rouge law enforcement. My family is  M86 heartbroken for the officers and their families. ...  M86 We reject violence of any kind directed at members of law  M86 enforcement or citizens.My hope is that one day soon we can come M86  together and find solutions to the very important issues  M86 facing our nation rather than continuing to hurt one M86  another.President Barack Obama on Sunday condemned the  M86 killings and all attacks on law enforcement.We as a  M86 nation have to be loud and clear that nothing justifies violence  M86 against law enforcement, Obama said, speaking from the  M86 White House press briefing room. Attacks on police are an M86  attack on all of us and the rule of law that makes society  M86 possible. In a written statement earlier in the day, Obama M86  called the Baton Rough shootings a cowardly and reprehensible  M86 assault.'No talking, just shooting'The shooting Sunday M86  took place around 8:40 a.m. (9:40 a.m. ET) in the city of M86  about 230,000 people, already tense after the high-profile M86  police shooting of Sterling, an African-American man, on July  M86 5.On Sunday, police received a call of a suspicious person  M86 walking down Airline Highway with an assault rifle, a source  M86 with knowledge of the investigation told CNN.When  M86 police arrived, the shooting began.There was no talking,  M86 just shooting, Baton Rouge Police Cpl. L.J. McKneely said.At  M86 an afternoon news conference, local and state authorities, M86  including Louisiana Gov. John", Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightLongestWord() {

        Word.run(function (context) {

            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // variable for keeping the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {

                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = context.document.body.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');

                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync)
        })
        .catch(errorHandler);
    }


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
    //this is to support getelementbyid().remove() function
    Element.prototype.remove = function () {
        this.parentElement.removeChild(this);
    }
    NodeList.prototype.remove = HTMLCollection.prototype.remove = function () {
        for (var i = this.length - 1; i >= 0; i--) {
            if (this[i] && this[i].parentElement) {
                this[i].parentElement.removeChild(this[i]);
            }
        }
    }

})();