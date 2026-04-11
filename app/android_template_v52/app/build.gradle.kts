import java.io.FileInputStream
import java.util.Properties

plugins {
    id("com.android.application")
    id("kotlin-android")
    id("dev.flutter.flutter-gradle-plugin")
}

val keystoreProperties = Properties()
val keystorePropertiesFile = rootProject.file("key.properties")
if (keystorePropertiesFile.exists()) {
    keystoreProperties.load(FileInputStream(keystorePropertiesFile))
}

fun signingValue(propertyKey: String, envKey: String): String? {
    return (keystoreProperties[propertyKey] as String?)?.takeIf { it.isNotBlank() }
        ?: System.getenv(envKey)?.takeIf { it.isNotBlank() }
}

val resolvedStoreFile = signingValue("storeFile", "CM_KEYSTORE_PATH")
val resolvedStorePassword = signingValue("storePassword", "CM_KEYSTORE_PASSWORD")
val resolvedKeyAlias = signingValue("keyAlias", "CM_KEY_ALIAS")
val resolvedKeyPassword = signingValue("keyPassword", "CM_KEY_PASSWORD")
val hasReleaseSigning = listOf(
    resolvedStoreFile,
    resolvedStorePassword,
    resolvedKeyAlias,
    resolvedKeyPassword,
).all { !it.isNullOrBlank() }

android {
    namespace = "com.grantproof"
    compileSdk = flutter.compileSdkVersion
    ndkVersion = flutter.ndkVersion

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = JavaVersion.VERSION_17.toString()
    }

    defaultConfig {
        applicationId = "com.grantproof"
        minSdk = 21
        targetSdk = flutter.targetSdkVersion
        versionCode = flutter.versionCode
        versionName = flutter.versionName
    }

    signingConfigs {
        create("release") {
            if (hasReleaseSigning) {
                storeFile = file(resolvedStoreFile!!)
                storePassword = resolvedStorePassword
                keyAlias = resolvedKeyAlias
                keyPassword = resolvedKeyPassword
            }
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            isShrinkResources = false
            signingConfig = if (hasReleaseSigning) {
                signingConfigs.getByName("release")
            } else {
                signingConfigs.getByName("debug")
            }
        }
    }
}

flutter {
    source = "../.."
}
