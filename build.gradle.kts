plugins {
    kotlin("jvm") version "1.9.23"
//    kotlin("plugin.serialization") version "1.9.23"
}

group = "org.mvk"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    // Apache POI for Excel file handling
    implementation("org.apache.poi:poi:5.2.3")
    implementation("org.apache.poi:poi-ooxml:5.2.3")

    // JDOM for XML parsing
    implementation("org.jdom:jdom2:2.0.6.1")

    testImplementation("org.jetbrains.kotlin:kotlin-test")

    implementation("org.jetbrains.kotlinx:dataframe:0.13.1")
    implementation("org.jetbrains.kotlinx:dataframe-excel:0.13.1")
//    implementation("org.jetbrains.kotlinx:kotlinx-serialization-json:1.6.0")
//    implementation("org.jetbrains.kotlinx:kotlinx-serialization-xml:1.5.0")
}

tasks.test {
    useJUnitPlatform()
}
kotlin {
    jvmToolchain(17)
}