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
		return false;
	},
	retrieveOverdueTasks: function ()
	{
		if (LZW.Xrm.ApplicationRibbon.OverdueTaskNotificationId === null)
		{
			Xrm.WebApi.retrieveMultipleRecords("task", "?$select=activityid&$filter=Microsoft.Dynamics.CRM.EqualUserId(PropertyName='ownerid') and Microsoft.Dynamics.CRM.LastXYears(PropertyName='scheduledend',PropertyValue=99) and statecode eq 0").then(

			function success(result)
			{
				if (result.entities.length > 0)
				{
					LZW.Xrm.ApplicationRibbon.showOverdueTaskNotification(result.entities.length);
				}
			},

			function (error)
			{
				Xrm.Navigation.openAlertDialog(
				{
					text: error.message
				});
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
				Xrm.Navigation.navigateTo(pageInput,
				{
					target: 1
				});
				Xrm.App.clearGlobalNotification(LZW.Xrm.ApplicationRibbon.OverdueTaskNotificationId);
			}
		};
		var notification = {
			type: 2,
			level: LZW.Xrm.ApplicationRibbon.NotificationLevel.Warning,
			message: `You have($
			{
				taskCount
			}) overdue task(s).`,
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
			Xrm.Navigation.openAlertDialog(
			{
				text: error.message
			});
		});
	}
};