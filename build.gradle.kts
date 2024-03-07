plugins {
	java
}

group = "com.example"
version = "0.0.1-SNAPSHOT"

java {
	sourceCompatibility = JavaVersion.VERSION_17
}

repositories {
	mavenCentral()
	maven {
		url = uri("https://releases.aspose.com/java/repo/")
	}
}

dependencies {
	implementation("com.aspose:aspose-cells:24.2")
}
