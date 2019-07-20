# Site-Down-Acker
A small Python app to track when to send out notifications or re-check status of sites and devices

In our Network Operations Center, we have to send out notifications at the one hour and three hour mark when a site is down. We also have to track when a Helmer refridgerator goes out of temperature specs or an Aruba wireless controller is rebooted and when it finally comes back online.
This app will allow you to track when those notifications need to go out or whem you need to check the status of an Aruba or Helmer. When it is time, an email notification is sent to a chosen email address thorugh your Outlook client.
After a notification event, the timer is reset to notify agian in fifteen minutes.
It keeps a display of all pending notifications in the order in which you will be notified with the time the notification will be sent out.
