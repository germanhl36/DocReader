//
//  testApp.swift
//  test
//
//  Created by German Huerta on 28/02/26.
//

import SwiftUI
import CoreData

@main
struct testApp: App {
    let persistenceController = PersistenceController.shared

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environment(\.managedObjectContext, persistenceController.container.viewContext)
        }
    }
}
