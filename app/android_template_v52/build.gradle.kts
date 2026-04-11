allprojects {
    repositories {
        google()
        mavenCentral()
    }
}

val sharedBuildDir = rootProject.layout.buildDirectory.dir("../../build").get()
rootProject.layout.buildDirectory.value(sharedBuildDir)

subprojects {
    val projectBuildDir = sharedBuildDir.dir(project.name)
    project.layout.buildDirectory.value(projectBuildDir)
}

subprojects {
    project.evaluationDependsOn(":app")
}

tasks.register<Delete>("clean") {
    delete(rootProject.layout.buildDirectory)
}
