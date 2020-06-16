var LZW = LZW || {};
LZW.Xrm = LZW.Xrm || {};
LZW.Xrm.ApplicationRibbon = {
	OverdueTaskNotificationId: null,
	NotificationLevel: {
		Success: 1,
		Error: 2,
		Warning: 3,
		Information: 4
	},
	GlobalNotificationEnableRule: function ()
	{
        this.retrieveOverdueTasks();
        this.showServiceUpdate("Dynamics 365 – Service Update 192 for OCE was deployed on 27 June 2020", "https://docs.microsoft.com/en-us/business-applications-release-notes/dynamics/released-versions/weekly-releases/update192");
        this.showServiceUpdate("Dynamics 365 – Service Update 189 for OCE was deployed on 20 June 2020", "https://docs.microsoft.com/en-us/business-applications-release-notes/dynamics/released-versions/weekly-releases/update189");
        this.showServiceUpdate("Dynamics 365 – Service Update 186 for OCE was deployed on 13 June 2020", "https://docs.microsoft.com/en-us/business-applications-release-notes/dynamics/released-versions/weekly-releases/update186");
		return false;
	},
	retrieveOverdueTasks: function ()
	{
		if (LZW.Xrm.ApplicationRibbon.OverdueTaskNotificationId === null)
		{
            Xrm.WebApi.retrieveMultipleRecords("task", "?$select=activityid\
            &$filter=Microsoft.Dynamics.CRM.EqualUserId(PropertyName='ownerid') \
            and Microsoft.Dynamics.CRM.LastXYears(PropertyName='scheduledend',PropertyValue=99) and statecode eq 0").then(
			function success(result)
			{
				if (result.entities.length > 0)
				{
					LZW.Xrm.ApplicationRibbon.showOverdueTaskNotification(result.entities.length);
				}
			},
			function (error)
			{
				Xrm.Navigation.openAlertDialog({ text: error.message });
			});
		}
	},
	showOverdueTaskNotification: function (taskCount)
	{
		var viewTaskAction = {
			actionLabel: "View My Tasks",
			eventHandler: function ()
			{
				var pageInput = {
					pageType: "entitylist",
					entityName: "task",
					viewId: "6cf285aa-eb20-4277-925a-3e9735411ff0" // My Tasks View
				};
				Xrm.Navigation.navigateTo(pageInput, { target: 1 });
				Xrm.App.clearGlobalNotification(LZW.Xrm.ApplicationRibbon.OverdueTaskNotificationId);
			}
		};
		var notification = {
			type: 2,
			level: LZW.Xrm.ApplicationRibbon.NotificationLevel.Warning,
			message: `You have(${taskCount}) overdue task(s).`,
			showCloseButton: true,
			action: viewTaskAction
		};
		Xrm.App.addGlobalNotification(notification).then(
		function success(result)
		{
			LZW.Xrm.ApplicationRibbon.OverdueTaskNotificationId = result;
		},
		function (error)
		{
			Xrm.Navigation.openAlertDialog({ text: error.message });
		});
    },
    showServiceUpdate: function (message, url)
	{
        if (localStorage.getItem("AnnouncementSeen"))
        { 
            return; 
        }
        var viewAdditionalInformation = {
			actionLabel: "Additional Information",
            eventHandler: function ()
            {
                localStorage.setItem("AnnouncementSeen", true);
                Xrm.Navigation.openUrl(url)
            }
		};
		var notification = {
			type: 2,
			level: LZW.Xrm.ApplicationRibbon.NotificationLevel.Information,
			message: message,
			showCloseButton: false,
			action: viewAdditionalInformation
		};
		Xrm.App.addGlobalNotification(notification).then(
		function success(result)
		{
			setTimeout(function(){ Xrm.App.clearGlobalNotification(result); }, 10000);
		},
		function (error)
		{
			Xrm.Navigation.openAlertDialog({ text: error.message });
		});
    }
};
