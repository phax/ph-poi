# ph-poi

[![javadoc](https://javadoc.io/badge2/com.helger/ph-poi/javadoc.svg)](https://javadoc.io/doc/com.helger/ph-poi)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.helger/ph-poi/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.helger/ph-poi) 

Java library with some Apache POI improvements. Also adds some helper functions to more easily read and write type safe Excel files.

# Maven usage

Add the following to your pom.xml to use this artifact, replacing `x.y.z` with the real version:

```xml
<dependency>
  <groupId>com.helger</groupId>
  <artifactId>ph-poi</artifactId>
  <version>x.y.z</version>
</dependency>
```

# News and noteworthy

* v5.3.2 - 2022-03-06
    * Updated to POI 5.2.1
* v5.3.1 - 2022-02-01
    * Updated to POI 5.2.0
* v5.3.0 - 2021-11-05
    * Updated to POI 5.1.0
    * Removed the class `POISLF4JLogger`
* v5.2.0 - 2021-03-21
    * Updated to ph-commons 10
* v5.1.0 - 2021-01-28
    * Updated to POI 5.0.0
    * Excluded support for SVG and PDF to lower dependency weight
* v5.0.7 - 2020-02-17
    * Updated to commons-compress 1.20
    * Updated to POI 4.1.2
* v5.0.6 - 2019-10-21
    * Updated to commons-compress 1.19
    * Updated to POI 4.1.1
* v5.0.5 - 2019-06-04
    * Updated to XMLBeans 3.1.0
    * Updated to POI 4.1.0
* v5.0.4 - 2019-02-09
    * Updated to POI 4.0.1
    * `WorkbookCreationHelper` now implements `AutoClosable`
* v5.0.3 - 2018-11-22
    * Updated to ph-commons 9.2.0
* v5.0.2 - 2018-09-10
    * Updated to POI 4.0.0
* v5.0.1 - 2018-08-09
    * Fixed OSGI ServiceProvider configuration
    * Added new class `CExcel` with some constants.
    * Updated to XMLBeans 3.0.0
* v5.0.0 - 2017-11-06
    * Updated to ph-commons 9.0.0
    * Removed deprecated methods
    * Updated to POI 3.17
* v4.1.1 - 2017-04-19
    * Updated to POI 3.16
* v4.1.0 - 2016-09-22
    * Updated to POI 3.15
* v4.0.0 - 2016-06-10
    * Requires at least JDK 8
    * Updated to POI 3.14
    * Binds to ph-commons 8.x
* v3.0.1 - 2015-10-19
* v3.0.0 - 2015-07-09
    * Binds to ph-commons 6.x
* v2.9.4 - 2015-03-31
* v2.9.3 - 2015-03-11
* v2.9.2 - 2015-01-13
* v2.9.1 - 2014-10-30
* v2.9.0 - 2014-08-25   

---

My personal [Coding Styleguide](https://github.com/phax/meta/blob/master/CodingStyleguide.md) |
On Twitter: <a href="https://twitter.com/philiphelger">@philiphelger</a> |
Kindly supported by [YourKit Java Profiler](https://www.yourkit.com)