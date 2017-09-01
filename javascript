# alteryx-sharepoint
<!DOCTYPE html>
<html lang="en">
<head>
	<title>Capacity Forecast Tool</title>
	<meta charset="UTF-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge;chrome=1" />
	<!--CSS-->
	<style>
		.btn-bs-file {
			position: relative;
		}
		.btn{
			display: block;
		}
		.btn-lg, .btn-group-lg>.btn {
			padding: 10px 16px;
			font-size: 18px;
			line-height: 1.33;
			border-radius: 6px;
		}
		.validationMessage {
			color: #f00;
		}
		#myBar {
			width: 1%;
			height: 10px;
			background-color: #5bc0de;
		}
		.loader {
			border: 16px solid #f3f3f3;
			border-radius: 50%;
			border-top: 16px solid #3498db;
			width: 75px;
			height: 75px;
			-webkit-animation: spin 2s linear infinite;
			animation: spin 2s linear infinite;
		}
		@-webkit-keyframes spin {
			0% { -webkit-transform: rotate(0deg); }
			100% { -webkit-transform: rotate(360deg); }
		}
		@keyframes spin {
			0% { transform: rotate(0deg); }
			100% { transform: rotate(360deg); }
		}
		li{
			list-style: none;
		}
		
	</style>
	<link rel="Stylesheet" href="https://sites.sprint.com/network/tpsd/SiteAssets/css/default.css" />
	<link href="https://sites.sprint.com/network/tpsd/SiteAssets/css/bootstrap.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/select2/4.0.3/css/select2.min.css">
	<link rel="stylesheet" href="https://sites.sprint.com/network/tpsd/SiteAssets/css/select2-bootstrap.css">

</head>   
<body>
<div class="codeRunner" style="padding:5px;">
	<div id="form" data-bind="with:ScenarioListViewModel">
	<div class="container1">
		<div class="row">
			<div class="col-sm-2">
				<h4 class="header-title m-t-0 m-b-30">Select Version:</h4>
				<select id="dropdown" required class="form-control select2" data-bind="options: release_cycles,value:selectedCycle"></select>
			</div>
			<div class="col-sm-3" data-bind="with:RefreshDatabaseViewModel">
				<a href="#"><h4 onclick="move()" data-bind="click:executeRefreshDatabaseWorkflow" class="header-title m-t-0 m-b-30" style="padding-top:35px;">Refresh List of Databases</h4></a>
				<!-- Modal to show the status of Refresh database workflow execution-->
				<div id="RefreshDatabaseModal" class="modal fade" role="dialog">
				<div class="modal-dialog">
					<!-- Modal content-->
					<div class="modal-content">
						<div class="modal-header">
							<h4 class="modal-title">Status</h4>
						</div>
						<div class="modal-body">
							<p><strong><span data-bind="text:jobStatus()"><div id="myBar"></div></span></strong></p>
						</div>
					</div>
				</div>
			</div>
			</div>
		</div>
		<div class="row" style="padding-top:10px;">
			<div class="col-sm-4">
				<h4 class="header-title m-t-0 m-b-30">Select Scenario:</h4>
				<select id="dropdown2" required="required" class="form-control select2" data-bind="options: scenarios, optionsText:'scenarioName',optionsCaption:'Please Select Scenario', value:selectedScenario,validationOptions: { errorElementClass: 
                                                 'input-validation-error' },selectedOptions: chosenScenario">
				</select>
			</div>
		</div>
		<div class="row" style="padding-top:10px;">
			<div class="col-sm-2" data-bind="enable:chosenScenario">
				<h4 class="header-title m-t-0 m-b-30">B41 Split Mode</h4>
				<input type="radio" data-bind="checked: splitMode" name="option" value="-SplitMode true"> True
  				<input type="radio" data-bind="checked: splitMode" name="option" value="-SplitMode false"> False
			</div>
			<div class="col-sm-2" data-bind="enable:chosenScenario,visible:disableFDmimo">
				<h4 class="header-title m-t-0 m-b-30">Apply FD-MIMO:</h4>
				<input  id="radio1"  type="radio" name="option1"> True
  				<input  id="radio2"  type="radio" name="option1"> False
			</div>
			<div class="col-sm-2" data-bind="visible: currentUserIsAdmin(),enable:chosenScenario">
				<h4 class="header-title m-t-0 m-b-30" style="color:#696969">Full Output (admin only)</h4>
				<input  class="full" data-bind="checked: fullOutput" type="radio" name="option2" value="-FullOutput true"> True
  				<input  class="full" data-bind="checked: fullOutput" type="radio" name="option2" value="-FullOutput false"> False
			</div>
		</div>
		<!--<div class="row" style="padding-top:30px;" data-bind="">-->
		<div class="row" style="padding-top:30px;" data-bind="visible:disableBandInputs">
			<div class="col-sm-3">
				<h4 class="header-title m-t-0 m-b-30">Override B25 Spectrum Timeline:</h4>
				<input  id="B25Input1" data-bind="checked: B25Input" type="radio" name="option3" value="true"> True
  				<input  id="B25Input2" data-bind="checked: B25Input" type="radio" name="option3" value="false"> False
            </div>
		</div>
		<div class="row" style="padding-top:30px;" data-bind="visible:showBrowseOptions1">
			<div class="col-sm-2"  data-bind="with:SpectrumPlanningViewModel">
				<label class="btn-bs-file btn btn-lg btn-primary">Browse excel file
                <input type="file" style="display:none;" name="files" id="inputFile1" data-bind=" event:{change:loadSheetNamesB25}" />
				</label>
				<span data-bind="text: fileNameXl1"></span>
			</div>
			<div class="col-sm-2"  data-bind="with:SpectrumPlanningViewModel">
				<select id="dropdown3" required="required" class="form-control select2" data-bind="options:xlnames,value:selectedSheetname1, 
                                                 selectedOptions: chosenFile,optionsCaption:'Please Select Sheet Name'">
				</select>
			</div>
		</div>
		<!--<div class="row" style="padding-top:30px;" data-bind="">-->
		<div class="row" style="padding-top:30px;" data-bind="visible:disableBandInputs">
			<div class="col-sm-3">
				<h4 class="header-title m-t-0 m-b-30">Override B41 Spectrum Timeline:</h4>
				<input  id="B41Input1" data-bind="checked: B41Input" type="radio" name="option4" value="true"> True
  				<input  id="B41Input2" data-bind="checked: B41Input" type="radio" name="option4" value="false"> False
			</div>
		</div>
		<div class="row" style="padding-top:30px;" data-bind="visible:showBrowseOptions2">
			<div class="col-sm-2" data-bind="with:SpectrumPlanningViewModel">
				<label class="btn-bs-file btn btn-lg btn-primary">Browse excel file
                <input type="file" style="display:none;" name="files" id="inputFile2"  data-bind=" event:{change:loadSheetNamesB41}"/>
				</label>
				<span data-bind="text: fileNameXl2"></span>
			</div>
			<div class="col-sm-2"  data-bind="with:SpectrumPlanningViewModel">
				<select id="dropdown4" required="required" class="form-control select2" data-bind="options:xlnames1,value:selectedSheetname2,
                                                 selectedOptions: chosenFile,optionsCaption:'Please Select Sheet Name(Macro)'">
				</select>
			</div>
			<div class="col-sm-2"  data-bind="with:SpectrumPlanningViewModel">
				<select id="dropdown5" required="required" class="form-control select2" data-bind="options:xlnames2,value:selectedSheetname3,
                                                 selectedOptions: chosenFile,optionsCaption:'Please Select Sheet Name(MiniMacro)'">
				</select>
			</div>
		</div>
		<div class="row" style="padding-top:30px;">
			<div class="col-sm-2">
				<button type="button" class="btn btn-info btn-lg" data-bind="click:validate">Execute Model</button>
			</div>
			<div class="col-sm-3">
				<button type="button" class="btn btn-info btn-lg" data-bind="with:CapacityPlanningViewModel,click:SearchHistoryIfJobExists">Check History</button>
			</div>
            
			<!-- Modal to show all the selections-->
			<div id="ConfirmSelectionModal" class="modal fade" role="dialog">
				<div class="modal-dialog">
					<!-- Modal content-->
					<div class="modal-content">
						<div class="modal-header">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h4 class="modal-title">Your Selection</h4>
						</div>
						<div class="modal-body">
							<p><strong>Database Name:</strong><span data-bind="text:db_name()"></p>
							<p><strong>Jar File Name:</strong>&nbsp;<span data-bind="text:jarFile()"></p>
							<p><strong>B41 Split Mode:</strong>&nbsp;<span data-bind="text:splitMode"></span></p>
							<p><strong>Apply FD-MIMO:</strong>&nbsp;<span data-bind="text:fdMimo()"></span></p>
							<p data-bind="visible: currentUserIsAdmin()"><strong>Full Output:</strong>&nbsp;<span data-bind="text:fullOutput"></span></p>
							<p data-bind="with:SpectrumPlanningViewModel"><strong>B25 Input:</strong>&nbsp;<span data-bind="text: fileNameXl1()"></span></p>
							<p data-bind="with:SpectrumPlanningViewModel"><strong>Sheet Name:</strong>&nbsp;<span data-bind="text:selectedSheetname1()"></span></p>
							<p data-bind="with:SpectrumPlanningViewModel"><strong>B41 Input:</strong>&nbsp;<span data-bind="text: fileNameXl2()"></span></p>
							<p data-bind="with:SpectrumPlanningViewModel"><strong>Macro:</strong>&nbsp;<span data-bind="text:selectedSheetname2()"></span></p>
							<p data-bind="with:SpectrumPlanningViewModel"><strong>MiniMacro:</strong>&nbsp;<span data-bind="text:selectedSheetname3()"></span></p>
						</div>
						<div class="modal-footer">
							<button type="button" data-bind="with:CapacityPlanningViewModel,click:onClickOk" class="btn btn-info btn-lg" data-dismiss="modal">Ok</button>
							<button type="button" class="btn btn-info btn-lg" data-dismiss="modal">Cancel</button>
						</div>
					</div>
				</div>
			</div>
			<form style="" id="execute" data-bind="ScenarioListViewModel">
			<!--<form style="display:none;" id="execute" data-bind="ScenarioListViewModel">-->
				<table>
					<tbody>
						<tr>
							<td class="name"><label>db name</label></td>
							<td><input name="db name" class="QuestionTextBox" type="text" data-bind="value:db_name()"></td>
						</tr>
						<tr>
							<td class="name"><label>SplitMode</label></td>
							<td><input name="SplitMode" class="QuestionListBox" type="text" data-bind="value:splitMode()"></td>
						</tr>
						<tr>
							<td class="name"><label>AdminOnly_FullOutput</label></td>
							<td><input name="AdminOnly_FullOutput" class="QuestionListBox" type="text" data-bind="value:fullOutput()"></td>
						</tr>
						<tr>
							<td class="name"><label>FD MIMO Sites</label></td>
							<td><input name="FD MIMO Sites" class="QuestionTextBox" type="text" data-bind="value:fdMimo()"></td>
						</tr>
						<tr>
							<td class="name"><label>JAR File</label></td>
							<td><input name="JAR File" class="QuestionTextBox" type="text" data-bind="value:jarFile()"></td>
						</tr>
						<tr>
							<td class="name"><label>B25 Spectrum Scenario</label></td>
							<td><input name="B25 Spectrum Scenario" class="QuestionTextBox" type="text" data-bind="value:excelSheetname()"></td>
						</tr>
						<tr>
							<td class="name"><label>B41 Spectrum Scenario (Macro)</label></td>
							<td><input name="B41 Spectrum Scenario (Macro)" class="QuestionTextBox" type="text" data-bind="value:excelSheetname1()"></td>
						</tr>
						<tr>
							<td class="name"><label>B41 Spectrum Scenario (MiniMacro)</label></td>
							<td><input name="B41 Spectrum Scenario (MiniMacro)" class="QuestionTextBox" type="text" data-bind="value:excelSheetname2()"></td>
						</tr>
					</tbody>
				</table>
				<!--<button type="button" class="btn btn-info btn-lg" data-bind="click:get" >Ok</button>-->
			</form>
            <!-- Modal for Validation-->
			<div id="ValidationModal" class="modal fade" role="dialog">
				<div class="modal-dialog">
					<!-- Modal content-->
					<div class="modal-content">
						<div class="modal-body">
							<p><h4><strong>Please check your Submission</strong></h4></p>
						</div>
						<div class="modal-footer">
							<button type="button" class="btn btn-info btn-lg" data-dismiss="modal">Ok</button>
						</div>
					</div>
				</div>
			</div>
			<div id="CapacityModal" class="modal fade" data-bind="with:CapacityPlanningViewModel" role="dialog" style="text-align:center;margin:auto;">
				<div class="modal-dialog">
					<!-- Modal content-->
					<div class="modal-content">
						<div class="modal-body">
							<div id="status-div">
								<div class="loader" style="text-align:center;margin:auto;"></div>
								<p><strong><span data-bind="text:jobStatus()"></span></strong></p>
							</div>
							<div id="errorMessages">
								
							</div>
						</div>
					</div>
				</div>
			</div>
            <div id = "resultDetails" style="padding:30px;" data-bind="with:CapacityPlanningViewModel">

			</div>
			<div id="showResults" style="padding:30px;" data-bind="with:CapacityPlanningViewModel">

			</div>
			<div id="outputDiv" style = "display:none;padding:30px;" data-bind="with:CapacityPlanningViewModel">
				<br>
				<a href="#"><h4 class="header-title m-t-0 m-b-30" id = 'download'>Download Zip file</h4></a>
				<br><br>
			</div>
		</div>
	</div>
	</div>
</div>
<!--Scripts-->
<script type="text/javascript" src="https://SiteAssets/js/knockout-3.4.2.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/o.min.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/bootstrap.min.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/knockout.validation.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/oauth-signature.min.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/Alteryx_api/alteryxGalleryAPI.js"></script>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.0/js/select2.full.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/xlsx.full.min.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/shim.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/jszip.js"></script>
<script type="text/javascript" src="https://SiteAssets/js/moment.js"></script>
<script type="text/javascript">
	function move() {
		var elem = document.getElementById("myBar");   
		var width = 1;
		var id = setInterval(frame, 100);
		function frame() {
			if (width >= 100) {
				clearInterval(id);
			} else {
				width++; 
				elem.style.width = width + '%'; 
			}
		}
	}
</script>
<script type="text/javascript">
	$( "#dropdown" ).select2({
		theme: "bootstrap"
	});
	$( "#dropdown2" ).select2({
		theme: "bootstrap"
	});
	$("#dropdown3").select2({
		theme: "bootstrap",
	});
	$("#dropdown4").select2({
		theme: "bootstrap",
	});
	$("#dropdown5").select2({
		theme: "bootstrap",
	});
</script>

<script>
	
	

</script>
<!------------------------------------------------------------------
 --    MVVM functions 
-------------------------------------------------------------------->

<script type="text/javascript">

	var sharepointBaseUrl = "https://sites.sprint.com/network/tpsd/trp/cp/";
	var sharepointScenarioListUrl = sharepointBaseUrl + "_api/lists/getbytitle('WebApp_Scenarios')";
	var adminGroupName = "TP&SD - Technology & Roadmap Planning Members";
	var workFlowName = "CapacityPlanning - Update SharePoint Metadata";
	var CapacityWorkflowName = "CapacityPlanning - Forecast Runner_v4";
	var historyList_Admin = "WebApp_ExecutionHistory";
	var historyList_Spectrum ="WebApp_ExecutionHistory_Spectrum";
	var historyList = "";
    var gallery = null;
    var ApiUrl = "";
    var ApiKey = "";
    var ApiSecret = "";
	var releaseCycleUrl;
	var scenarioNameUrl1;
	var scenarioNameUrl2;
	var url;
	var splitModeValue;
	var fullOutputValue;
	var fdMimoAvailable;
	var jobStat;
	var myvar;
	var displayName;
	var interval;
	var errorString;
	var X = XLSX;
	var XW = {
		/* worker message */
		msg: 'xlsx',
		/* worker scripts */
		rABS: './xlsxworker2.js',
		norABS: './xlsxworker1.js',
		noxfer: './xlsxworker.js'
	};
	var global_wb;
	var sheetNames;
	var xlf1 = document.getElementById('inputFile1');
	var xlf2 = document.getElementById('inputFile2');
	var xlFileName;
	var xlFileName1;
	var fileAndSheetName1;
	var fileAndSheetName2;
	var fileAndSheetName3;
	
	

/*------Refresh database--------------------------------------------------------------------------------------------------------------------------------------------*/
	//model for refresh database workflow
	function Workflow(alteryxWorkflow) {
        var self = this;
        self.id = alteryxWorkflow.id;
        self.name = alteryxWorkflow.metaInfo.name;
        // self.tagText = ko.computed( function() { return self.name + " (" + self.id + ")"; }, self);
    }
	//viewmodel for refresh database workflow
	var RefreshDatabaseViewModel= function ()  {
        var self = this;
        
		self.Workflows = ko.observableArray([]);
		self.jobStatus = ko.observable();
		self.executeRefreshDatabaseWorkflow = function() {
            self.getWorkflowsFromAlteryx();
        }
        //get all workflows from alteryx
        self.getWorkflowsFromAlteryx = function(){
            self.Workflows([]);
            gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
			gallery.getSubscriptionWorkflows( function(workflows){
                for (var i = 0, len = workflows.length; i<len; i++){
					// console.log(new Workflow(workflows[i]));
					if(workflows[i].metaInfo.name == workFlowName ){
						var idWorkFlow = workflows[i].id;
                        self.getQuestions(idWorkFlow);
                        break;
                    }
                }
            }, function(response){
                var error = response.responseJSON && response.responseJSON.message || response.statusText;
                console.log("workflow not found - Error in function getWorkflowsFromAlteryx");
            })
        };
        //Get questions for Refresh database Workflow
        self.getQuestions = function(idWorkFlow){
            gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
            gallery.getAppQuestions(idWorkFlow, function(questions){
                var questions1 = questions;
                self.executeRefreshWorkflow(idWorkFlow,questions1);
            },function(response){
                var error = response.responseJSON && response.responseJSON.message || response.statusText;
                console.log("Questions not found - Error in function getQuestions");
            })
        };
        //Execute Refresh database workflow
        self.executeRefreshWorkflow = function(idWorkFlow,questions1){
            gallery.executeWorkflow(idWorkFlow, questions1, function(job){
                var jobId = job.id;
                //get job details
                var getStatus = gallery.getJob(jobId, function(job){
                    // alert(job.status + " - " + job.disposition);
                    self.jobStatus("Running");
                    $("#RefreshDatabaseModal").modal({
                        backdrop: 'static',
                        keyboard: false
                    });
                    setTimeout(function(){ location.reload(); }, 10000);
                }, function(response){
                        var error = response.responseJSON && response.responseJSON.message || response.statusText;
                        alert("Job status cannot be found");
                })				
            }, function(response){
                var error = response.responseJSON && response.responseJSON.message || response.statusText;
                console.log("Workflow cannot be executed - Error in function executeRefreshWorkflow");
            })
        };
    };
/*------Scenarios--------------------------------------------------------------------------------------------------------------------------------------------*/
	//model for scenarios
	function ScenarioModel(sharepointItem){
		var self = this;
		self.scenarioName = ko.observable(sharepointItem.ScenarioName);
		self.release_cycle = ko.observable(sharepointItem.Release_x0020_Cycle);
		self.title = ko.observable(sharepointItem.Title);
		self.jar = ko.observable(sharepointItem.JAR_x0020_File);
		self.fdmimo = ko.observable(sharepointItem.FDMIMO_x0020_Candidate_x0020_Sit);
        self.availableForSpectrum = ko.observable(sharepointItem.Available_x0020_for_x0020_Spectr);
		self.statusActiveDisabled = ko.observable(sharepointItem.Status);
	}
	
	//viewModel for scenarios
	var ScenarioListViewModel= function (){
		//Data
		var self = this;
		self.scenarios = ko.observableArray([]);
		self.release_cycles= ko.observableArray([]);
		self.currentUserIsAdmin = ko.observable(false);
		self.selectedCycle = ko.observable();
		self.fullOutput = ko.observable("-FullOutput false");
		self.splitMode = ko.observable("-SplitMode false");
		self.fdMimo = ko.observable();
		self.db_name = ko.observable();
		self.jarFile = ko.observable();
		self.chosenScenario=ko.observable();
		self.disableFDmimo = ko.observable();
		self.errorMessages = ko.observable();
		self.selectedScenario = ko.observable();
		self.excelSheetname = ko.observable("");
		self.excelSheetname1 = ko.observable("");
		self.excelSheetname2 = ko.observable("");
		self.disableBandInputs = ko.observable();
		self.showBrowseOptions1 = ko.observable(false);
		self.showBrowseOptions2 = ko.observable(false);
		self.B25Input = ko.observable("false");
		self.B41Input = ko.observable("false");
		//get user-group info
		$.ajax( {
			url: sharepointBaseUrl + "_api/web/currentUser/groups?$select=title",
			method: "GET",
			headers: { "Accept": "application/json; odata=verbose" },
			success: function (data) {
				var items = data.d.results;
				// console.log(items);
				for (var i = 0, len = items.length; i < len; i++) {
					if (items[i].Title == adminGroupName) { 
						self.currentUserIsAdmin(true); 
						
						historyList = historyList_Admin;
						releaseCycleUrl = "/items?$filter=Status eq 'Active'&$orderby=Release_x0020_Cycle desc";
						scenarioNameUrl1 = "/items?$select=Title,ScenarioName,JAR_x0020_File,Available_x0020_for_x0020_Spectr,FDMIMO_x0020_Candidate_x0020_Sit&$filter=(Release_x0020_Cycle eq ";
						scenarioNameUrl2 = ") and (Status eq 'Active')";
						break;
					}
					else{
					
						historyList = historyList_Spectrum;
						releaseCycleUrl = "/items?$filter=Status eq 'Active' and Available_x0020_for_x0020_Spectr eq '1'&$orderby=Release_x0020_Cycle desc";
						scenarioNameUrl1 = "/items?$select=Title,ScenarioName,JAR_x0020_File,Available_x0020_for_x0020_Spectr,FDMIMO_x0020_Candidate_x0020_Sit&$filter=(Release_x0020_Cycle eq ";
						scenarioNameUrl2 = ") and (Status eq 'Active') and (Available_x0020_for_x0020_Spectr eq '1')";
						
					}
				}

				self.pushReleaseCycles(releaseCycleUrl,scenarioNameUrl1,scenarioNameUrl2);
				self.excelSheetname("");
				self.excelSheetname1("");
				self.excelSheetname2("");
				
				$("#B25Input1").click(function(){self.showBrowseOptions1(true);});
				$("#B25Input2").click(function(){self.showBrowseOptions1(false);self.excelSheetname("");$("#dropdown3").empty();});
				$("#B41Input1").click(function(){self.showBrowseOptions2(true);$("#inputFile2").val("");});
				$("#B41Input2").click(function(){self.showBrowseOptions2(false);self.excelSheetname1("");self.excelSheetname2("")});
			},
			error: function (data) { console.log("ERROR in fetching User group Info: " + data);console.log(data);}
		});

		//load release cycles in first dropdown
		self.pushReleaseCycles = function(releaseCycleUrl,scenarioNameUrl1,scenarioNameUrl2){
			
			$.ajax({
				url: sharepointScenarioListUrl + releaseCycleUrl,
				method: "GET",
				headers: { "Accept": "application/json; odata=verbose" },
				success: function (data) {
					var items = data.d.results;
					var defaultChosen = false;

					items.forEach( function(item) { 
						var choice = item.Release_x0020_Cycle;

						if (self.release_cycles.indexOf(choice) == -1) { 
							self.release_cycles.push(choice);

							if (!defaultChosen) { 
								//set Scenarios for default Release selection
								self.pushScenariosToDropdown(choice,scenarioNameUrl1,scenarioNameUrl2);
								defaultChosen=true;								
							}
						}
					});
					
					self.selectedCycle.subscribe(function(value) {
						self.showBrowseOptions1(false);
						self.showBrowseOptions2(false);
						self.scenarios([]);
						$("#inputFile1").empty();
						$("#inputFile2").empty();
						self.fdMimo("");
						self.excelSheetname("");
						self.excelSheetname1("");
						self.excelSheetname2("");
						self.pushScenariosToDropdown(value,scenarioNameUrl1,scenarioNameUrl2);
					});
				},
				error: function (data) { console.log("ERROR in function pushReleaseCycles : " + data);console.log(data);}
			});
		};
		
		//load scenarios in second dropdown
		self.pushScenariosToDropdown = function(value,scenarioNameUrl1,scenarioNameUrl2){
			self.B25Input = ko.observable("false");
			self.B41Input = ko.observable("false");
			
			$.ajax( {
				url: sharepointScenarioListUrl + scenarioNameUrl1 + value + scenarioNameUrl2,
				method: "GET",
				headers: { "Accept": "application/json; odata=verbose" },
				success: function (data) {
					var items = data.d.results;
					items.forEach( function(item) { 
						self.scenarios.push(new ScenarioModel(item));console.log(data);
					});
					self.selectedScenario.subscribe(function(value) {
						if(value !== ''){
							if (value !== undefined){
								var dbName = ko.toJSON(value.title);
								var jarFile1 = ko.toJSON(value.jar);
								var fdMimoAvailable = ko.toJSON(value.fdmimo);
								var availableToSpectrum = ko.toJSON(value.availableForSpectrum);
								if(availableToSpectrum == "true"){
									self.disableBandInputs(true);
								}
								if(availableToSpectrum == "false"){
									
									self.disableBandInputs(false);
									self.showBrowseOptions1(false);
									self.showBrowseOptions2(false);
									self.excelSheetname("");
									self.excelSheetname1("");
									self.excelSheetname2("");
								}
								self.setValues(dbName,jarFile1,fdMimoAvailable); 
							}
						}
					});
				},
				error: function (data) { console.log("ERROR in function pushScenariosToDropdown: " + data);console.log(data);}
			});
		};
		
		//pass selected values to modal
		self.setValues = function(dbName,jarFile1,fdMimoAvailable){
			$("#radio2").prop('checked', true);
			if(fdMimoAvailable =="null"){
				self.fdMimo("");
				self.disableFDmimo(false);
			} else{
				self.disableFDmimo(true);
			}
			self.db_name(dbName.replace (/(^")|("$)/g, ''));
			self.jarFile(jarFile1.replace (/(^")|("$)/g, ''));
			$("#radio1").click(function(){self.fdMimo(fdMimoAvailable.replace (/(^")|("$)/g, ''));});
			$("#radio2").click(function(){self.fdMimo("");});
		};
		
		//Validate form
		self.selectedScenario = ko.observable().extend({required: true});
		self.errors = ko.validation.group(ScenarioListViewModel);
		self.validate= function(){
			if (self.errors().length === 0) {
				$("#ConfirmSelectionModal").modal();
			}else {
				$("#ValidationModal").modal();
				self.errors.showAllMessages();
			}
		};
		self.errors = ko.validation.group(self);
	};
/*-----------------Capacity Planning Workflow-------------------------------------------------------------------------------------------------------------------*/
		
	var CapacityPlanningViewModel= function (){
		self.results = ko.observableArray([]);
		self.jobStatus = ko.observable();
		
		//get Logged in user's name
		$.ajax( {
			url: sharepointBaseUrl + "_api/SP.UserProfiles.PeopleManager/GetMyProperties",
			method: "GET",
			headers: { "Accept": "application/json; odata=verbose" },
			success: function (data) {
				displayName = data.d.DisplayName;
			},
			error: function (data) { console.log("ERROR in fetching logged in user's name");}
		});	
		//On Clicking Ok button
		self.onClickOk = function(){
			if($('#inputFile1').val()){
				var fileInput = $('#inputFile1');
				self.AddFile(fileInput);
			}
			if($('#inputFile2').val()){
				var fileInput = $('#inputFile2');
				self.AddFile(fileInput);
			}
			self.jobStatus("Running");
			$("#CapacityModal").modal();
			$("#status-div").show();
			$("#errorMessages").hide();
            self.getAllWorkflowsFromAlteryx();
		};
		self.AddFile = function(fileInput){
			
			var fileName = fileInput[0].files[0].name;
			var reader = new FileReader();
			reader.onload = function (e) {
			var fileData = e.target.result;
				$.ajax({
					url: sharepointBaseUrl + "/_api/web/GetFolderByServerRelativeUrl('/network/tpsd/trp/cp/Spectrum Planning Scenarios')/Files/add(url='" + fileName + "',overwrite=true)",
					method: "POST",
					binaryStringRequestBody: true,
					data: fileData,
					processData: false,
					headers: {
						"ACCEPT": "application/json;odata=verbose",  
						"X-RequestDigest": myvar,
						"content-length": fileData.byteLength
					},                                                                                                                            
					success: function (data) {       
						//alert("Uploaded successfully");      
					},
					error: function (data) {
						alert("Error occured." + data.responseText);
						console.log(data);
					}
				});                 
			};
			reader.readAsArrayBuffer(fileInput[0].files[0]);
		};
		// get all executions from Execution history Sharepoint list
		self.SearchHistoryIfJobExists = function (){
			$.ajax( {
				url: sharepointBaseUrl + "_api/web/lists/GetByTitle('"+ historyList+"')/items?$select=Title,DB_Name,JAR_File,Split_Mode,Full_Output,Disposition,Workflow_Status,FD_MIMO,Date,Sheet_Name_B25,Sheet_Name_B41_Macro,Sheet_Name_B41_Mini_Macro",
				method: "GET",
				headers: { "Accept": "application/json; odata=verbose" },
				success: function (data) {
					var itemsSortedByDate;
					var items = data.d.results;
					console.log("Before sort:");
					console.log(items);
					itemsSortedByDate = items.sort(function(a,b) { 
						return new Date(b.Date).getTime() - new Date(a.Date).getTime();
					});
					
					console.log("After sort:");
					console.log(itemsSortedByDate);
					self.getJobIdForCapacity(itemsSortedByDate);
				},
				error: function (data) { console.log(data);console.log("ERROR in function SearchHistoryIfJobExists: " + data) }
			});
		};
		self.getJobIdForCapacity = function(itemsSortedByDate){
			var jobIdForCapacity;
			var found = false;
			for (var i = 0, len = itemsSortedByDate.length; i < len; i++) {
				var db = itemsSortedByDate[i].DB_Name;
				var split_mode = itemsSortedByDate[i].Split_Mode;
				var full_output = itemsSortedByDate[i].Full_Output;
				var workkflow_status = itemsSortedByDate[i].Workflow_Status;
				var disposition1 = itemsSortedByDate[i].Disposition;
				var fd;
				var B25;
				var B41Macro;
				var B41MiniMacro;
				var fdFromForm;
				var B25FromForm;
				var B41MacroFromForm;
				var B41MiniMacroFromForm;
				if(ScenarioListViewModel.fdMimo() == undefined){
					fdFromForm = "";
				}
				else{
					fdFromForm = ScenarioListViewModel.fdMimo();
				}
				if(ScenarioListViewModel.excelSheetname() == undefined){
					B25FromForm = "";
				}
				else{
					B25FromForm = ScenarioListViewModel.excelSheetname();
				}
				if(ScenarioListViewModel.excelSheetname1() == undefined){
					B41MacroFromForm = "";
				}
				else{
					B41MacroFromForm = ScenarioListViewModel.excelSheetname1();
				}
				if(ScenarioListViewModel.excelSheetname2() == undefined){
					B41MiniMacroFromForm = "";
				}
				else{
					B41MiniMacroFromForm = ScenarioListViewModel.excelSheetname2();
				}
				if(itemsSortedByDate[i].FD_MIMO === null){
					 fd = "";
					 console.log("fd from sharepoint list if null"+ fd);
					 console.log("fd from form if null"+ ScenarioListViewModel.fdMimo());
				}
				else{
					 fd = itemsSortedByDate[i].FD_MIMO;
					 console.log(ScenarioListViewModel.fdMimo());
				}
				if(itemsSortedByDate[i].Sheet_Name_B25 === null){
					 B25 = "";
					 console.log("B25 from sharepoint list if null:" +B25);
					 console.log("file from form if null:" + ScenarioListViewModel.excelSheetname());
				}
				else{
					 B25 = itemsSortedByDate[i].Sheet_Name_B25;
					 console.log("B25 if not null:" +B25);
				}
				if(itemsSortedByDate[i].Sheet_Name_B41_Macro === null){
					 B41Macro = "";
					 console.log("B41Macro from sharepoint list if null:" +B41Macro);
					 console.log("file from form:" + ScenarioListViewModel.excelSheetname1());
				}
				else{
					 B41Macro = itemsSortedByDate[i].Sheet_Name_B41_Macro;
					 console.log("B41Macro if not null:" + B41Macro);
				}
				if(itemsSortedByDate[i].Sheet_Name_B41_Mini_Macro === null){
					 B41MiniMacro = "";
					 console.log("B41MiniMacro from sharepoint list if null:" + B41MiniMacro);
					 console.log("file from form:" + ScenarioListViewModel.excelSheetname2());
				}
				else{
					 B41MiniMacro = itemsSortedByDate[i].Sheet_Name_B41_Mini_Macro;
					 console.log("B41MiniMacro if not null:" + B41MiniMacro);
				}
				console.log("Database:" + db);
				console.log("database from form:" + ScenarioListViewModel.db_name());
				console.log("SplitMode:" + split_mode);
				console.log("SplitMode from form:" + ScenarioListViewModel.splitMode());
				console.log("Full output:" + full_output);
				console.log("Full output from form:" + ScenarioListViewModel.fullOutput());
				console.log("Status from sharepoint:" + workkflow_status);
				console.log("Disposition from Sharepoint:" + disposition1);
				
				
				if((db == ScenarioListViewModel.db_name())
				&& (split_mode  == ScenarioListViewModel.splitMode() )
				&& (full_output == ScenarioListViewModel.fullOutput() )
				&& (workkflow_status ===  "Running" || "Queued" || "Completed") 
				&& (disposition1 === "Success" || "None") 
				&& (fd === fdFromForm)
				&& (B25 ===  B25FromForm)
				&& (B41Macro ===  B41MacroFromForm)
				&& (B41MiniMacro === B41MiniMacroFromForm)){
					jobIdForCapacity = itemsSortedByDate[i].Title;
					console.log("Job Id:" + jobIdForCapacity);
					console.log("fdmimo from form: " + ScenarioListViewModel.fdMimo());
					console.log("fdmimo from list: " + fd);
					console.log("B25 from form: " + fileAndSheetName1);
					console.log("B25 from list: " + B25);
					console.log("B41Macro from form: " + fileAndSheetName2);
					console.log("B41Macro from list: " + B41Macro);
					console.log("B41MiniMacro from form: " + fileAndSheetName3);
					console.log("B41MiniMacro from list: " + B41MiniMacro);
					self.getJobResults(jobIdForCapacity);
					found = true;
					break;
				}
			}
			if (!found) {
  				alert("Job not found in Sharepoint Execution History List. Click Execute Model to run");
				$("#showResults").empty();
				$("#resultDetails").empty();
				$("#outputDiv").hide();
			}	
		};
		self.getJobResults = function(jobIdForCapacity){
			gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
			gallery.getJob(jobIdForCapacity, function(job){
                var jobStat = job.status;
				var jobDisposition = job.disposition;
				var jobCreateDate = job.createDate;
				var outputs = job.outputs;
				var messages = job.messages;
				self.displayMessages(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,outputs, messages);
				//get list id from execution History Sharepoint list and update Execution History list
				self.getListIdFromHistory(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,messages);
			},function(response){
				var error = response.responseJSON && response.responseJSON.message || response.statusText;
				console.log(error);
				alert("Job has been deleted from Alteryx Execution History. Click Execute Model to run");
			});
		};
		//get all workflows
		self.getAllWorkflowsFromAlteryx = function (){
			gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
			gallery.getSubscriptionWorkflows( function(workflows){
				console.log("All Workflows from");
				console.log(workflows);
				workflowsSortedByDate = workflows.sort(function(a,b) { 
					return new Date(b.uploadDate).getTime() - new Date(a.uploadDate).getTime() 
				});
				console.log("All Workflows after sorting by date");
				console.log(workflowsSortedByDate);
				for (var i=0, len=workflowsSortedByDate.length; i<len; i++){
					if(workflows[i].metaInfo.name == CapacityWorkflowName){
						idCapacityWorkFlow = workflows[i].id;
						console.log("Workflow Id:" + idCapacityWorkFlow);
						self.executeCapacityWorkflow(idCapacityWorkFlow);
                        break;
					}
				}
			}, function(response){
				var error = response.responseJSON && response.responseJSON.message || response.statusText;
				console.log("Error in function getAllWorkflowsFromAlteryx" + error);
			});
		};
		//Execute Capacity Workflow
		self.executeCapacityWorkflow = function(idCapacityWorkFlow){
			gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
			var questionsForCapacity = $('form').serializeArray();
			gallery.executeWorkflow(idCapacityWorkFlow, questionsForCapacity, function(job){
				var jobIdForCapacity = job.id;
				console.log("Newly Created Job Id for Capacity" + jobIdForCapacity);
				interval = setInterval(function(){self.getJobDetails(jobIdForCapacity)}, 3000), runSaveHistoryfunctionOnce=true;
				
			},function(response){
				var error = response.responseJSON && response.responseJSON.message || response.statusText;
				console.log("Error in function executeCapacityWorkflow" + error);
			});
		};
		//get job details
		self.getJobDetails = function (jobIdForCapacity){
            gallery = new Gallery(ApiUrl.trim(), ApiKey.trim(), ApiSecret.trim());
			gallery.getJob(jobIdForCapacity, function(job){
                var jobStat = job.status;
				var jobDisposition = job.disposition;
				var jobCreateDate = job.createDate;
				
				var outputs = job.outputs;
				var messages = job.messages;
                //run Save History function only once as it has three job status - Running , Queued and Completed and hence will be saved thrice in the list
				if (runSaveHistoryfunctionOnce) { 
            		self.saveHistoryToList(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate);
					runSaveHistoryfunctionOnce = false; 
				}
				if(jobStat=="Completed"){
					clearInterval(interval);
					
					//display success/error message on modal
					self.displayMessages(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,outputs, messages);
                    
                    //get list id from execution History Sharepoint list and update Execution History list
					self.getListIdFromHistory(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,messages);
				}
			},function(response){
				var error = response.responseJSON && response.responseJSON.message || response.statusText;
				console.log("Error in function getJobDetails/getJob" + error);
			});	
		};
		self.displayMessages = function(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,outputs, messages){
			var errormessages =$("#errorMessages");
			var showOutput = $("#showResults");
			var resultDetails = $("#resultDetails");
			$("#status-div").hide();
			errormessages.show();
			
			if(jobDisposition=="Success"){
				var src="https://sites.sprint.com/network/tpsd/SiteAssets/images/check.png"
				errormessages.html("<img src="+ src +" height='42' width='42'>");
			}
			else if(jobDisposition=="Error"){
				var src="https://sites.sprint.com/network/tpsd/SiteAssets/images/error.png"
				errormessages.html("<img src="+ src +" height='42' width='42'>");
			}
			errormessages.append("<li>Workflow executed " + " - " + jobStat + " - " + jobDisposition + " on " + new Date(jobCreateDate).toLocaleString() +  "</li>");
			$("#resultDetails").show();
			resultDetails.html("<br><br><br><br><li>Workflow executed " + " - " + jobStat + " - " + jobDisposition + " on " + new Date(jobCreateDate).toLocaleString() + "</li><br>");
			//if there are any errors display error messages
			var message;
			
			var errorString;
			for (var j = 0; j < messages.length; j++) {
				message = messages[j];
				if (message.status === 3){
					errorString += "<li>";
					errorString += message.text += (message.toolId > 0) ? " (Tool Id: " + message.toolId + ")" : "";
					errorString += "</li>";
				}
			}
			errormessages.append(errorString);
			//get output
			self.getOutput(outputs,jobIdForCapacity);
		};
		//get Output in the form of html or zip file
        self.getOutput = function(outputs,jobIdForCapacity){
			console.log("Output files:");
			console.log(outputs);
			var outputId;
			var outputIdForZip;
            for (var i = 0; i < outputs.length; i++){
				if(outputs[i].formats.length == 6){
					outputId = outputs[i].id;
					break;
				}
			}
			for (var i = 0; i < outputs.length; i++){
				if(outputs[i].name == "Forecast.zip"){
					outputIdForZip = outputs[i].id;
					break;
				}
			}
			console.log("Output Id for Html:" + outputId);
			console.log("Output Id for Zip:" + outputIdForZip);
            var format = "Html";
            var urlForHtml = gallery.getOutputFileURL(jobIdForCapacity, outputId, format);
            //get html file
           	self.loadResult(urlForHtml);
			$("#download").click(function() {
				var formatAsZip = "Raw";
				var urlForZip = gallery.getOutputFileURL(jobIdForCapacity, outputIdForZip, formatAsZip);
				window.location.assign(urlForZip);
			});
		};
		//load result inline in html format
		self.loadResult = function(urlForHtml){
			 $.ajax({
                type: "GET",
                url: urlForHtml,
                contentType: "application/html; charset=utf-8",
                contentDisposition: "inline",
                success: function(result){
					$("#showResults").html(result);
					$("#outputDiv").show();
					var resultDiv = document.getElementById("resultDetails").innerHTML;
					var showResultsDiv = document.getElementById("showResults").innerHTML;
					var outputDiv = document.getElementById("outputDiv").innerHTML;
					var htmlContent = [resultDiv + showResultsDiv + outputDiv];
					var bl = new Blob(htmlContent, {type: "text/html"});
					var a = document.createElement("a");
					a.href = URL.createObjectURL(bl);
					console.log(a.href);
                },
                error: function(){
                    alert("Unable to load Html and Zip file");
                }
            });
		};
		//get form digest value
		$.ajax({
			url: sharepointBaseUrl + "/_api/contextinfo",
			type: "POST",
			headers: {
				"accept": "application/json;odata=verbose",
				"contentType": "text/xml"
			},
			success: function (data) {
				myvar = data.d.GetContextWebInformation.FormDigestValue;
			},
			error: function (data) { console.log("ERROR in fetching form digest value: " + data.responseText) }
		});
		//add workflow to list
		self.saveHistoryToList = function(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate){
			var date = new Date(jobCreateDate);
			var dateTime = moment(date).format("YYYY/MM/DD HH:mm:ss");
			var requestData = ko.toJSON({ '__metadata': { 'type': 'SP.Data.WebApp_x005f_ExecutionHistoryListItem' }, 'Title':jobIdForCapacity,'DB_Name':ScenarioListViewModel.db_name(), 
			'Split_Mode':ScenarioListViewModel.splitMode(),'Full_Output':ScenarioListViewModel.fullOutput(),
			'JAR_File':ScenarioListViewModel.jarFile(),'User':displayName,'Workflow_Status':jobStat, 
			'Disposition':jobDisposition,'Date':new Date(jobCreateDate), 'FD_MIMO':ScenarioListViewModel.fdMimo(),
			'Local_Date': dateTime,
			'Sheet_Name_B25':fileAndSheetName1,
			'Sheet_Name_B41_Macro':fileAndSheetName2,'Sheet_Name_B41_Mini_Macro':fileAndSheetName3});
			$.ajax({
				url: sharepointBaseUrl + "_api/web/lists/GetByTitle('WebApp_ExecutionHistory')/items",
				method: "POST",
				headers: {     
					"accept": "application/json;odata=verbose",
					"content-type": "application/json;odata=verbose",
					"content-length": requestData.length,
					"X-RequestDigest":myvar,
				},
				data: requestData,
				success: function (data) {  },
				error: function (data) { console.log("ERROR : " + data.responseText + requestData);console.log("ERROR in function saveHistoryToList ");}
			});
		};
		//get the list item id from history list 
		self.getListIdFromHistory = function(jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,messages){
            $.ajax({
                url: sharepointBaseUrl + "_api/web/lists/GetByTitle('"+ historyList+"')/items?$select=Id&$filter=Title+eq+'"+jobIdForCapacity+"'", 
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: function (data) {
                    var items = data.d.results;
					var idfromlist;
                    for (var i = 0; i < items.length; i++){
						idfromlist = items[0].Id;
                    }
                    //update list with new job status
                   self.updateExecutionHistoryList(idfromlist,jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,messages);
                },
                error: function (data) { console.log("ERROR: " + data);console.log("ERROR in function getListIdFromHistory "); }
            });
			
		};
		//update list with new job status -completed/Error/Success
		self.updateExecutionHistoryList = function(idfromlist,jobIdForCapacity,jobStat,jobDisposition,jobCreateDate,messages){
			self.geterrorString(messages);
			var date = new Date(jobCreateDate);
			var dateTime = moment(date).format("YYYY/MM/DD HH:mm:ss");
			var requestData = ko.toJSON({ '__metadata': { 'type': 'SP.Data.WebApp_x005f_ExecutionHistoryListItem' },'Workflow_Status':jobStat, 'Disposition':jobDisposition,'Date':new Date(jobCreateDate),'Errors':errorString,'Local_Date': dateTime});
			$.ajax({
				url: sharepointBaseUrl + "_api/web/lists/GetByTitle('WebApp_ExecutionHistory')/items(" + idfromlist + ")",
				method: "POST",
				headers: {     
					"X-HTTP-Method":"MERGE",
					"accept": "application/json;odata=verbose",
					"content-type": "application/json;odata=verbose",
					"content-length": requestData.length,
					"X-RequestDigest": myvar,
					"IF-MATCH":"*"
				},
				data: requestData,
				success: function (data) {  },
				error: function (data) { 
					console.log("ERROR in function updateExecutionHistoryList" );
					console.log("ERROR in function updateExecutionHistoryList" + data.responseText);
					console.log("Requestdata: " + requestData);
				}
			});
		};
		self.geterrorString = function(messages){
			var message;
			
			for (var j = 0; j < messages.length; j++) {
				message = messages[j];
				if (message.status === 3){
					errorString += message.text += (message.toolId > 0) ? " (Tool Id: " + message.toolId + ")" : "";
				}
			}
			console.log("error string: "+ errorString);
			console.log("Messages");
			console.log(messages);
		};
		
	};

	//model for scenarios
	function SpectrumModel(sharepointItem1){
	var self = this;
		self.excelFileName = ko.observable(sharepointItem1.Name);
	};

	var SpectrumPlanningViewModel= function (){
		self.excelFiles= ko.observableArray([]);
		self.selectedSheetname1 = ko.observable("");
		self.selectedSheetname2 = ko.observable("");
		self.selectedSheetname3 = ko.observable("");
		self.chosenFile = ko.observable();
		self.xlnames = ko.observableArray([]);
		self.xlnames1 = ko.observableArray([]);
		self.xlnames2 = ko.observableArray([]);
		self.fileNameXl1 = ko.observable();
		self.fileNameXl2 = ko.observable();
		
		self.handleFile = function(vm, evt) {
			var files = evt.target.files;
			var f = files[0];
			{
				var reader = new FileReader();
				xlFileName = f.name;
				self.fileNameXl1(xlFileName);
				console.log(xlFileName);
				reader.onload = function(e) {
					if(typeof console !== 'undefined') console.log("onload", new Date());
					var data = e.target.result;
					var wb;
					var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
					process_wb(wb);
				}
			}
			reader.readAsArrayBuffer(f);
		};
		self.handleFile1 = function(vm, evt) {
			var files = evt.target.files;
			var f = files[0];
			{
				var reader = new FileReader();
				xlFileName1 = f.name;
				self.fileNameXl2(xlFileName1);
				console.log(xlFileName1);
				reader.onload = function(e) {
					if(typeof console !== 'undefined') console.log("onload", new Date());
					var data = e.target.result;
					var wb;
					var arr = fixdata1(data);
					wb = X.read(btoa(arr), {type: 'base64'});
					process_wb1(wb);
				}
			}
			reader.readAsArrayBuffer(f);
		};
		self.fixdata = function(data) {
			var o = "", l = 0, w = 10240;
			for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
			o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
			return o;
		};
		self.fixdata1 = function(data) {
			var o = "", l = 0, w = 10240;
			for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
			o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
			return o;
		};
		self.process_wb = function(wb) {
			self.xlnames("");
			var output = "";
			output = JSON.stringify(to_json(wb), 2, 2);
			var parse = JSON.parse(output);
			sheetNames = Object.keys(parse);
			console.log("this is output");
			console.log(sheetNames);
			// // $('#inputFile1').change(self.xlnames(sheetNames));
			// // $('#inputFile2').change(self.xlnames1(sheetNames),self.xlnames2(sheetNames));
			// var abc = $('#inputFile1').val();
			// var xyz = $('#inputFile2').val();	
			// if () {
				self.xlnames(sheetNames);
				// self.xlnames2(sheetNames);
			// } 
			
			if(typeof console !== 'undefined') console.log("output", new Date());
		};
		self.process_wb1 = function(wb) {
			self.xlnames1("");
			self.xlnames1("");
			var output = "";
			output = JSON.stringify(to_json1(wb), 2, 2);
			var parse = JSON.parse(output);
			sheetNames = Object.keys(parse);
			console.log("this is output");
			console.log(sheetNames);
			// // $('#inputFile1').change(self.xlnames(sheetNames));
			// // $('#inputFile2').change(self.xlnames1(sheetNames),self.xlnames2(sheetNames));
			// var abc = $('#inputFile1').val();
			// var xyz = $('#inputFile2').val();	
			// if () {
				self.xlnames1(sheetNames);
				self.xlnames2(sheetNames);
			// } 
			
			if(typeof console !== 'undefined') console.log("output", new Date());
		};
		self.to_json = function(workbook) {
			var result = {};
			workbook.SheetNames.forEach(function(sheetName) {
				var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
				if(roa.length > 0){
					result[sheetName] = roa;
				}
			});
			return result;
		};
		self.to_json1 = function(workbook) {
			var result = {};
			workbook.SheetNames.forEach(function(sheetName) {
				var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
				if(roa.length > 0){
					result[sheetName] = roa;
				}
			});
			return result;
		};
		self.loadSheetNamesB25 = function(vm, evt){
			self.xlnames([]);
			self.handleFile(vm, evt);
		};
		self.loadSheetNamesB41 = function(vm, evt){
			self.xlnames1([]);
			self.xlnames2([]);
			self.handleFile1(vm, evt);
		};
		self.selectedSheetname1.subscribe(function(newValue) {
			ScenarioListViewModel.excelSheetname("");
			if($('#B25Input1').prop('checked')){
				if(newValue !== undefined){
					if (newValue !== undefined){
						fileAndSheetName1 =xlFileName + "|||`" + newValue + "$`"
						ScenarioListViewModel.excelSheetname(xlFileName + "|||`" + newValue + "$`");
					}
				}	
			}
			
		});
		self.selectedSheetname2.subscribe(function(newValue) {
			ScenarioListViewModel.excelSheetname1("");
			if(newValue !== ''){
				if (newValue !== undefined){
					fileAndSheetName2 = xlFileName1 + "|||`" + newValue + "$`"
					ScenarioListViewModel.excelSheetname1(xlFileName1 + "|||`" + newValue + "$`");
				}
			}	
		});
		self.selectedSheetname3.subscribe(function(newValue) {
			ScenarioListViewModel.excelSheetname2("");
			if(newValue !== ''){
				if (newValue !== undefined){
					fileAndSheetName3 = xlFileName1 + "|||`" + newValue + "$`";
					ScenarioListViewModel.excelSheetname2(xlFileName1 + "|||`" + newValue + "$`");
				}
			}	
		});
	}

	ko.validation.init({
		registerExtenders: true,
		messagesOnModified: true,
		insertMessages: true,
		parseInputAttributes: false,
		messageTemplate: null,
		decorateInputElement : true
	}, true);
	//Master ViewModel
	var masterVM = (function(){
		this.ScenarioListViewModel =  new ScenarioListViewModel(),
		this.RefreshDatabaseViewModel = new RefreshDatabaseViewModel(this.ScenarioListViewModel);
		this.CapacityPlanningViewModel = new CapacityPlanningViewModel(this.ScenarioListViewModel);
		this.SpectrumPlanningViewModel =  new SpectrumPlanningViewModel(this.ScenarioListViewModel);
		
	})();
	ko.applyBindings(masterVM);
	
</script>
</body>
</html>
