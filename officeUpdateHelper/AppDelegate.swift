/*

  AppDelegate.swift
  officeUpdateHelper

  Created by Gavin Pardoe on 03/03/2016.
  Updated 15/03/2016.

  Designed for Use with JAMF Casper Suite.

  The MIT License (MIT)

  Copyright (c) 2016 Gavin Pardoe


Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/


import Cocoa
import WebKit

@NSApplicationMain
class AppDelegate: NSObject, NSApplicationDelegate, NSUserNotificationCenterDelegate {

    @IBOutlet weak var window: NSWindow!
    @IBOutlet weak var webView: WebView!
    @IBOutlet weak var activityWheel: NSProgressIndicator!
    @IBOutlet weak var label: NSTextField!
    @IBOutlet weak var deferButtonDisable: NSButton!
    @IBOutlet weak var installButtonDisable: NSButton!

    var appCheckTimer = NSTimer()

    func applicationDidFinishLaunching(aNotification: NSNotification) {
        
        let htmlPage = NSBundle.mainBundle().URLForResource("index", withExtension: "html")
        webView.mainFrame.loadRequest(NSURLRequest(URL: htmlPage!))

        activityWheel.displayedWhenStopped = false
        self.window.makeKeyAndOrderFront(nil)
        self.window.level = Int(CGWindowLevelForKey(.FloatingWindowLevelKey))
        self.window.level = Int(CGWindowLevelForKey(.MaximumWindowLevelKey))
        
        label.allowsEditingTextAttributes = true
        label.stringValue = "Checking for Running Office Apps"
        installButtonDisable.enabled = false
        NSUserNotificationCenter.defaultUserNotificationCenter().delegate = self
        appCheck()
        
    }

    func applicationWillTerminate(aNotification: NSNotification) {
        // Insert code here to tear down your application
    }
    
    
    func userNotificationCenter(center: NSUserNotificationCenter, shouldPresentNotification notification: NSUserNotification) -> Bool {
        return true
    }
    
    
    func appCheck() {
        
        appCheckTimer = NSTimer.scheduledTimerWithTimeInterval(4.0, target: self, selector: #selector(AppDelegate.officeAppsRunning), userInfo: nil, repeats: true)
        
    }
    
    
    func officeAppsRunning() {
        
        for runningApps in NSWorkspace.sharedWorkspace().runningApplications {
            if let _ = runningApps.localizedName {
                
            }
        }
        
        
        let appName = NSWorkspace.sharedWorkspace().runningApplications.flatMap { $0.localizedName }
        
        if appName.contains("Microsoft Word") {
            
            label.textColor = NSColor.redColor()
            label.stringValue = "Please Quit All Office Apps to Continue!"
            
        } else if appName.contains("Microsoft Excel") {
            
            label.textColor = NSColor.redColor()
            label.stringValue = "Please Quit All Office Apps to Continue!"
            
        } else if appName.contains("Microsoft PowerPoint") {
            
            label.textColor = NSColor.redColor()
            label.stringValue = "Please Quit All Office Apps to Continue!"
            
        } else if appName.contains("Microsoft Outlook") {
            
            label.textColor = NSColor.redColor()
            label.stringValue = "Please Quit All Office Apps to Continue!"
            
        } else if appName.contains("Microsoft OneNote") {
            
            label.textColor = NSColor.redColor()
            label.stringValue = "Please Quit All Office Apps to Continue!"
            
        } else {
            
            installUpdates()
        }
        
    }
    
    
    func installUpdates() {
        
        appCheckTimer.invalidate()
        label.textColor = NSColor.blackColor()
        label.stringValue = "Click Install to Begin Updating Office"
        installButtonDisable.enabled = true
        
    }
    
    @IBAction func installButton(sender: AnyObject) {
        
        installButtonDisable.enabled = false
        deferButtonDisable.enabled = false
        self.activityWheel.startAnimation(self)
        let notification = NSUserNotification()
        notification.title = "Office Updater"
        notification.informativeText = "Installing Updates..."
        notification.soundName = NSUserNotificationDefaultSoundName
        NSUserNotificationCenter.defaultUserNotificationCenter().deliverNotification(notification)
        label.stringValue = "Installing Updates, Window will Close Once Completed"
        NSAppleScript(source: "do shell script \"/usr/local/jamf/bin/jamf policy -event officeUpdate\" with administrator " + "privileges")!.executeAndReturnError(nil)
        
        // For Testing
        //let alert = NSAlert()
        //alert.messageText = "Testing Completed..!"
        //alert.addButtonWithTitle("OK")
        //alert.runModal()
        
    }
    
    @IBAction func deferButton(sender: AnyObject) {
        
        let notification = NSUserNotification()
        notification.title = "Office Updater"
        notification.informativeText = "Installation has Been Defered..."
        notification.soundName = NSUserNotificationDefaultSoundName
        NSUserNotificationCenter.defaultUserNotificationCenter().deliverNotification(notification)

        NSApplication.sharedApplication().terminate(self)
        
    }

}


