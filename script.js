'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheets = dashboard.worksheets;

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(function (parameters) {
            parameters.forEach(function (p) {
              p.addEventListener(tableau.TableauEventType.ParameterChanged, onParameterChange);
            });
          });

          function onParameterChange (parameterChangeEvent) {
            parameterChangeEvent.getParameterAsync().then(function (param) {
                console.log("Change parameter");
            });
          }

        

        function updateAndRefreshData() {
            //const dashboard = tableau.extensions.dashboardContent.dashboard;

            // Get the parameters P_FROM_DATE and P_TO_DATE
            dashboard.getParametersAsync().then(parameters => {
                const fromDateParam = parameters.find(param => param.name === 'P_FROM_DATE');
                const toDateParam = parameters.find(param => param.name === 'P_TO_DATE');
				const fdtdParam = parameters.find(param => param.name === 'P_FD_TD');
				
                if (!fromDateParam || !toDateParam) {
                    document.getElementById('status').innerText = "Missing parameters: P_FROM_DATE or P_TO_DATE.";
                    return;
                }

                const fromDate = fromDateParam.currentValue.value;
                const toDate = toDateParam.currentValue.value;
				const fromDateToDate = fromDate + ',' + toDate
				
				// Cập nhật P_FD bằng giá trị của P_FROM_DATE
				fdtdParam.changeValueAsync(fromDateToDate).then(() => {
                                console.log('thay doi tham so FD-TD');
                                
                            });
							
            }).catch(err => {
                document.getElementById('status').innerText = "Error fetching parameters.";
                console.error(err);
            });
        }
    });
});
