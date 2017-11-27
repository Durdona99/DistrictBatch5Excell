package com.districtbatch5.runners;

import org.junit.runner.RunWith;

import cucumber.api.CucumberOptions;
import cucumber.api.junit.Cucumber;

@RunWith(Cucumber.class)
@CucumberOptions(plugin = { "html:target/cucumber-reports",
		"json:target/jsonreport.json" }, features = "./src/test/resources/com/districtbatch5/features", glue = "com/districtbatch5/step_definitions", dryRun = false)
public class CukesRunner {

}
