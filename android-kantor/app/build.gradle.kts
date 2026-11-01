plugins {
    id("com.android.application")
}

android {
    namespace = "id.irvan.kantorvanx"
    compileSdk = 35

    defaultConfig {
        applicationId = "id.irvan.kantorvanx"
        minSdk = 23
        targetSdk = 35
        versionCode = 1
        versionName = "1.0.0"

        buildConfigField("String", "WEB_APP_URL", "\"https://irvanmaulana.my.id/kantor/\"")
    }

    buildFeatures {
        buildConfig = true
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(getDefaultProguardFile("proguard-android-optimize.txt"), "proguard-rules.pro")
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
}
