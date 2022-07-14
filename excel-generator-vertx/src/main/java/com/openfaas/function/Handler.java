package com.openfaas.function;

import io.vertx.ext.web.RoutingContext;
import io.vertx.core.json.JsonObject;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
// import org.joda.time.DateTime;
// import org.joda.time.DateTimeZone;
// import org.joda.time.format.DateTimeFormat;
// import org.joda.time.format.DateTimeFormatter;
import java.util.*;
import org.json.JSONObject;
import org.json.JSONArray;

import com.openfaas.function.RunFunctionResponse.java;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.License;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.SaveFormat;

import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.PutObjectRequest;
import software.amazon.awssdk.auth.credentials.AwsBasicCredentials;
import software.amazon.awssdk.auth.credentials.StaticCredentialsProvider;
import software.amazon.awssdk.services.s3.model.ListBucketsRequest;
import software.amazon.awssdk.services.s3.model.ListBucketsResponse;

public class Handler implements io.vertx.core.Handler<RoutingContext> {

	public RunFunctionResponse runFunctionResponse = new RunFunctionResponse();



	public void handle(RoutingContext routingContext) {
		
		byte[] buffer = null;
		java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
		
		
			String dummyResponse = "{\"dataMap\":{\"output\":[{\"regionCost\":4568.3,\"color\":\"#FF7F50\",\"latitude\":\"20.07\",\"selectable\":true,\"title\":\"Mumbai$4568.3\",\"region\":\"AsiaPacific(Mumbai)\",\"technicalName\":\"ap-south-1\",\"longitude\":\"73.87\"},{\"regionCost\":2144.17,\"color\":\"#65cea7\",\"latitude\":\"-33.865143\",\"selectable\":true,\"title\":\"Sydney$2144.17\",\"region\":\"AsiaPacific(Sydney)\",\"technicalName\":\"ap-southeast-2\",\"longitude\":\"151.209900\"},{\"regionCost\":1998.61,\"color\":\"#77787d\",\"latitude\":\"37.926868\",\"selectable\":true,\"title\":\"N.Virginia$1998.61\",\"region\":\"USEast(N.Virginia)\",\"technicalName\":\"us-east-1\",\"longitude\":\"-78.024902\"},{\"regionCost\":1747.29,\"color\":\"#6bafbd\",\"latitude\":\"50.108314\",\"selectable\":true,\"title\":\"Frankfurt$1747.29\",\"region\":\"EU(Frankfurt)\",\"technicalName\":\"eu-central-1\",\"longitude\":\"8.659025\"},{\"regionCost\":1552.51,\"color\":\"#A52A2A\",\"latitude\":\"1.290270\",\"selectable\":true,\"title\":\"Singapore$1552.51\",\"region\":\"AsiaPacific(Singapore)\",\"technicalName\":\"ap-southeast-1\",\"longitude\":\"103.851959\"},{\"regionCost\":1460.88,\"color\":\"#8A2BE2\",\"latitude\":\"35.706723\",\"selectable\":true,\"title\":\"Tokyo$1460.88\",\"region\":\"AsiaPacific(Tokyo)\",\"technicalName\":\"ap-northeast-1\",\"longitude\":\"139.423891\"},{\"regionCost\":712.58,\"color\":\"#DEB887\",\"latitude\":\"-39.13755099999999\",\"selectable\":true,\"title\":\"Miscellaneous$712.58\",\"region\":\"Miscellaneous\",\"longitude\":\"-137.98828125\"},{\"regionCost\":114.04,\"color\":\"#5F9EA0\",\"latitude\":\"44.447506\",\"selectable\":true,\"title\":\"Oregon$114.04\",\"region\":\"USWest(Oregon)\",\"technicalName\":\"us-west-2\",\"longitude\":\"-120.627888\"},{\"regionCost\":32.04,\"color\":\"#D2691E\",\"latitude\":\"14.775667\",\"selectable\":true,\"title\":\"AsiaPacific$32.04\",\"region\":\"AsiaPacific\",\"longitude\":\"105.174885\"},{\"regionCost\":29.19,\"color\":\"#6495ED\",\"latitude\":\"51.501116\",\"selectable\":true,\"title\":\"London$29.19\",\"region\":\"EU(London)\",\"technicalName\":\"eu-west-2\",\"longitude\":\"-0.134884\"},{\"regionCost\":3.13,\"color\":\"#DC143C\",\"latitude\":\"40.749081\",\"selectable\":true,\"title\":\"UnitedStates$3.13\",\"region\":\"UnitedStates\",\"longitude\":\"-102.031842\"},{\"regionCost\":2.62,\"color\":\"#00FFFF\",\"latitude\":\"24.343122\",\"selectable\":true,\"title\":\"India$2.62\",\"region\":\"India\",\"longitude\":\"78.832508\"},{\"regionCost\":2.09,\"color\":\"#00008B\",\"latitude\":\"37.532600\",\"selectable\":true,\"title\":\"Seoul$2.09\",\"region\":\"AsiaPacific(Seoul)\",\"technicalName\":\"ap-northeast-2\",\"longitude\":\"127.024612\"},{\"regionCost\":2.01,\"color\":\"#008B8B\",\"latitude\":\"-33.13755099999999\",\"selectable\":true,\"title\":\"Global$2.01\",\"region\":\"Any\",\"longitude\":\"-137.98828125\"},{\"regionCost\":1.05,\"color\":\"#B8860B\",\"latitude\":\"49.087998\",\"selectable\":true,\"title\":\"Europe$1.05\",\"region\":\"Europe\",\"longitude\":\"22.435091\"},{\"regionCost\":0.26,\"color\":\"#006400\",\"latitude\":\"26.0667\",\"selectable\":true,\"title\":\"MiddleEast(Bahrain)$0.26\",\"region\":\"MiddleEast(Bahrain)\",\"technicalName\":\"me-south-1\",\"longitude\":\"50.5577\"},{\"regionCost\":0.22,\"color\":\"#8B008B\",\"latitude\":\"40.001774\",\"selectable\":true,\"title\":\"Japan$0.22\",\"region\":\"Japan\",\"longitude\":\"140.773587\"},{\"regionCost\":0.08,\"color\":\"#556B2F\",\"latitude\":\"48.864716\",\"selectable\":true,\"title\":\"Paris$0.08\",\"region\":\"EU(Paris)\",\"technicalName\":\"eu-west-3\",\"longitude\":\"2.349014\"},{\"regionCost\":0.06,\"color\":\"#FF8C00\",\"latitude\":\"40.297415\",\"selectable\":true,\"title\":\"Ohio$0.06\",\"region\":\"USEast(Ohio)\",\"technicalName\":\"us-east-2\",\"longitude\":\"-82.964746\"},{\"regionCost\":0.05,\"color\":\"#8B0000\",\"latitude\":\"59.3293\",\"selectable\":true,\"title\":\"EU(Stockholm)$0.05\",\"region\":\"EU(Stockholm)\",\"technicalName\":\"eu-north-1\",\"longitude\":\"18.0686\"},{\"regionCost\":0.05,\"color\":\"#E9967A\",\"latitude\":\"53.065412\",\"selectable\":true,\"title\":\"Ireland$0.05\",\"region\":\"EU(Ireland)\",\"technicalName\":\"eu-west-1\",\"longitude\":\"-7.693532\"},{\"regionCost\":0.04,\"color\":\"#8FBC8F\",\"latitude\":\"22.3193\",\"selectable\":true,\"title\":\"HongKong$0.04\",\"region\":\"AsiaPacific(HongKong)\",\"technicalName\":\"ap-east-1\",\"longitude\":\"114.1694\"},{\"regionCost\":0.03,\"color\":\"#2F4F4F\",\"latitude\":\"60.1453591\",\"selectable\":true,\"title\":\"Canada(Central)$0.03\",\"region\":\"Canada(Central)\",\"technicalName\":\"ca-central-1\",\"longitude\":\"-110.6503695\"},{\"regionCost\":0.03,\"color\":\"#00CED1\",\"latitude\":\"40.778261\",\"selectable\":true,\"title\":\"N.California$0.03\",\"region\":\"USWest(N.California)\",\"technicalName\":\"us-west-1\",\"longitude\":\"-119.41793239999998\"},{\"regionCost\":0.02,\"color\":\"#FFD700\",\"latitude\":\"-23.55\",\"selectable\":true,\"title\":\"SaoPaulo$0.02\",\"region\":\"SouthAmerica(SaoPaulo)\",\"technicalName\":\"sa-east-1\",\"longitude\":\"-46.63\"},{\"regionCost\":0.01,\"color\":\"#3CB371\",\"latitude\":\"54.1776679\",\"selectable\":true,\"title\":\"Canada$0.01\",\"region\":\"Canada\",\"longitude\":\"-115.9652551\"},{\"regionCost\":0,\"color\":\"#808000\",\"latitude\":\"-14.297346\",\"selectable\":true,\"title\":\"SouthAmerica$0.0\",\"region\":\"SouthAmerica\",\"longitude\":\"-59.436785\"},{\"regionCost\":0,\"color\":\"#FFA500\",\"latitude\":\"34.672314\",\"selectable\":true,\"title\":\"Osaka$0.0\",\"region\":\"AsiaPacific(Osaka)\",\"technicalName\":\"ap-northeast-3\",\"longitude\":\"135.484802\"},{\"regionCost\":0,\"color\":\"#DA70D6\",\"latitude\":\"29.2985\",\"selectable\":true,\"title\":\"MiddleEast$0.0\",\"region\":\"MiddleEast\",\"longitude\":\"42.5510\"},{\"regionCost\":0,\"color\":\"#CD853F\",\"latitude\":\"-33.9249\",\"selectable\":true,\"title\":\"Africa(CapeTown)$0.0\",\"region\":\"Africa(CapeTown)\",\"longitude\":\"18.4241\"},{\"regionCost\":0,\"color\":\"#8B4513\",\"latitude\":\"45.4642\",\"selectable\":true,\"title\":\"EU(Milan)$0.0\",\"region\":\"EU(Milan)\",\"longitude\":\"9.1900\"},{\"regionCost\":0,\"color\":\"#9ACD32\",\"latitude\":\"-26.195246\",\"selectable\":true,\"title\":\"SouthAfrica$0.0\",\"region\":\"SouthAfrica\",\"longitude\":\"28.034088\"},{\"regionCost\":0,\"color\":\"#008080\",\"latitude\":\"-23.584581\",\"selectable\":true,\"title\":\"Australia$0.0\",\"region\":\"Australia\",\"longitude\":\"133.633786\"}],\"tableKeys\":[\"ServiceName\",\"AsiaPacific(Mumbai)\",\"AsiaPacific(Sydney)\",\"USEast(N.Virginia)\",\"EU(Frankfurt)\",\"AsiaPacific(Singapore)\",\"AsiaPacific(Tokyo)\",\"Miscellaneous\",\"USWest(Oregon)\",\"AsiaPacific\",\"EU(London)\",\"UnitedStates\",\"India\",\"AsiaPacific(Seoul)\",\"Any\",\"Europe\",\"MiddleEast(Bahrain)\",\"Japan\",\"EU(Paris)\",\"USEast(Ohio)\",\"EU(Stockholm)\",\"EU(Ireland)\",\"AsiaPacific(HongKong)\",\"Canada(Central)\",\"USWest(N.California)\",\"SouthAmerica(SaoPaulo)\",\"Canada\",\"SouthAmerica\",\"AsiaPacific(Osaka)\",\"MiddleEast\",\"Africa(CapeTown)\",\"EU(Milan)\",\"SouthAfrica\",\"Australia\",\"Total\"],\"table\":[{\"Africa(CapeTown)\":0,\"Total\":2.2,\"ServiceName\":\"AWSKeyManagementService\",\"AsiaPacific(Mumbai)\":0.25,\"AsiaPacific(Seoul)\":0.24,\"AsiaPacific(Singapore)\":0.01,\"AsiaPacific(Sydney)\":0.01,\"AsiaPacific(Tokyo)\":0.48,\"Canada(Central)\":0,\"EU(Frankfurt)\":0.01,\"EU(Ireland)\":0,\"EU(London)\":0,\"EU(Milan)\":0,\"EU(Paris)\":0,\"EU(Stockholm)\":0,\"MiddleEast(Bahrain)\":0,\"SouthAmerica(SaoPaulo)\":0,\"USEast(N.Virginia)\":1.19,\"USEast(Ohio)\":0,\"USWest(N.California)\":0,\"USWest(Oregon)\":0.01},{\"Africa(CapeTown)\":0,\"Total\":1.97,\"ServiceName\":\"AmazonSimpleQueueService\",\"AsiaPacific(Mumbai)\":1.26,\"AsiaPacific(Seoul)\":0,\"AsiaPacific(Singapore)\":0,\"AsiaPacific(Sydney)\":0,\"AsiaPacific(Tokyo)\":0,\"Canada(Central)\":0,\"EU(Frankfurt)\":0.22,\"EU(Ireland)\":0,\"EU(London)\":0.01,\"EU(Milan)\":0,\"EU(Paris)\":0,\"EU(Stockholm)\":0,\"MiddleEast(Bahrain)\":0,\"Miscellaneous\":0.19,\"SouthAmerica(SaoPaulo)\":0,\"USEast(N.Virginia)\":0.29,\"USEast(Ohio)\":0,\"USWest(N.California)\":0,\"USWest(Oregon)\":0},{\"Africa(CapeTown)\":0,\"Total\":42.22,\"ServiceName\":\"AWSCloudTrail\",\"AsiaPacific(HongKong)\":0,\"AsiaPacific(Mumbai)\":42.22,\"AsiaPacific(Osaka)\":0,\"AsiaPacific(Seoul)\":0,\"AsiaPacific(Singapore)\":0,\"AsiaPacific(Sydney)\":0,\"AsiaPacific(Tokyo)\":0,\"Canada(Central)\":0,\"EU(Frankfurt)\":0,\"EU(Ireland)\":0,\"EU(London)\":0,\"EU(Milan)\":0,\"EU(Paris)\":0,\"EU(Stockholm)\":0,\"MiddleEast(Bahrain)\":0,\"SouthAmerica(SaoPaulo)\":0,\"USEast(N.Virginia)\":0,\"USEast(Ohio)\":0,\"USWest(N.California)\":0,\"USWest(Oregon)\":0},{\"Africa(CapeTown)\":0,\"Total\":14371.29,\"ServiceName\":\"Total\",\"Any\":2.01,\"AsiaPacific\":32.04,\"AsiaPacific(HongKong)\":0.03,\"AsiaPacific(Mumbai)\":4568.29,\"AsiaPacific(Osaka)\":0,\"AsiaPacific(Seoul)\":2.09,\"AsiaPacific(Singapore)\":1552.5,\"AsiaPacific(Sydney)\":2144.16,\"AsiaPacific(Tokyo)\":1460.87,\"Australia\":0,\"Canada\":0.01,\"Canada(Central)\":0.03,\"EU(Frankfurt)\":1747.28,\"EU(Ireland)\":0.05,\"EU(London)\":29.19,\"EU(Milan)\":0,\"EU(Paris)\":0.08,\"EU(Stockholm)\":0.05,\"Europe\":1.05,\"India\":2.62,\"Japan\":0.22,\"MiddleEast\":0,\"MiddleEast(Bahrain)\":0.26,\"Miscellaneous\":712.58,\"SouthAfrica\":0,\"SouthAmerica\":0,\"SouthAmerica(SaoPaulo)\":0.02,\"USEast(N.Virginia)\":1998.61,\"USEast(Ohio)\":0.05,\"USWest(N.California)\":0.03,\"USWest(Oregon)\":114.04,\"UnitedStates\":3.13},{\"Africa(CapeTown)\":0,\"Total\":325.11,\"ServiceName\":\"AmazonCloudWatch\",\"Any\":1.84,\"AsiaPacific(HongKong)\":0.01,\"AsiaPacific(Mumbai)\":93.36,\"AsiaPacific(Seoul)\":0.05,\"AsiaPacific(Singapore)\":84.18,\"AsiaPacific(Sydney)\":48.27,\"AsiaPacific(Tokyo)\":0.54,\"Canada(Central)\":0.01,\"EU(Frankfurt)\":48.05,\"EU(Ireland)\":0.02,\"EU(London)\":0.1,\"EU(Paris)\":0.03,\"EU(Stockholm)\":0.02,\"MiddleEast(Bahrain)\":0.01,\"SouthAmerica(SaoPaulo)\":0,\"USEast(N.Virginia)\":48.28,\"USEast(Ohio)\":0.03,\"USWest(N.California)\":0.01,\"USWest(Oregon)\":0.3},{\"Africa(CapeTown)\":0,\"Total\":0,\"ServiceName\":\"AmazonSimpleNotificationService\",\"AsiaPacific(Mumbai)\":0,\"AsiaPacific(Seoul)\":0,\"AsiaPacific(Singapore)\":0,\"AsiaPacific(Sydney)\":0,\"AsiaPacific(Tokyo)\":0,\"Canada(Central)\":0,\"EU(Frankfurt)\":0,\"EU(Ireland)\":0,\"EU(London)\":0,\"EU(Milan)\":0,\"EU(Paris)\":0,\"EU(Stockholm)\":0,\"MiddleEast(Bahrain)\":0,\"Miscellaneous\":0,\"SouthAmerica(SaoPaulo)\":0,\"USEast(N.Virginia)\":0,\"USEast(Ohio)\":0,\"USWest(N.California)\":0,\"USWest(Oregon)\":0},{\"Any\":0.17,\"Total\":0.17,\"ServiceName\":\"SavingsPlansforAWSComputeusage\"},{\"Any\":0,\"Total\":0,\"ServiceName\":\"AWSSystemsManager\"},{\"Any\":0,\"Total\":0,\"ServiceName\":\"AmazonQuickSight\"},{\"AsiaPacific\":32.04,\"Total\":209.13,\"ServiceName\":\"AmazonCloudFront\",\"AsiaPacific(Mumbai)\":0,\"AsiaPacific(Singapore)\":0,\"AsiaPacific(Sydney)\":0,\"Australia\":0,\"Canada\":0.01,\"EU(Frankfurt)\":0,\"EU(Ireland)\":0,\"EU(London)\":0,\"Europe\":1.05,\"India\":2.62,\"Japan\":0.22,\"MiddleEast\":0,\"Miscellaneous\":170.06,\"SouthAfrica\":0,\"SouthAmerica\":0,\"USEast(N.Virginia)\":0,\"USEast(Ohio)\":0,\"USWest(N.California)\":0,\"USWest(Oregon)\":0,\"UnitedStates\":3.13},{\"AsiaPacific(HongKong)\":0.03,\"Total\":182.14,\"ServiceName\":\"AmazonSimpleStorageService\",\"AsiaPacific(Mumbai)\":80.16,\"AsiaPacific(Seoul)\":0.03,\"AsiaPacific(Singapore)\":7.9,\"AsiaPacific(Sydney)\":5.08,\"AsiaPacific(Tokyo)\":1.4,\"Canada(Central)\":0.01,\"EU(Frankfurt)\":20.43,\"EU(Ireland)\":0.02,\"EU(London)\":29.07,\"EU(Paris)\":0.05,\"EU(Stockholm)\":0.02,\"MiddleEast(Bahrain)\":0.01,\"Miscellaneous\":14.05,\"SouthAmerica(SaoPaulo)\":0.02,\"USEast(N.Virginia)\":17.37,\"USEast(Ohio)\":0.03,\"USWest(N.California)\":0.01,\"USWest(Oregon)\":6.45},{\"AsiaPacific(Mumbai)\":2323.4,\"Total\":9524.07,\"ServiceName\":\"AmazonElasticComputeCloud\",\"AsiaPacific(Seoul)\":1.77,\"AsiaPacific(Singapore)\":1276.96,\"AsiaPacific(Sydney)\":1718.62,\"AsiaPacific(Tokyo)\":1296.75,\"EU(Frankfurt)\":1319.14,\"Miscellaneous\":82.3,\"USEast(N.Virginia)\":1405.08,\"USWest(Oregon)\":100.05},{\"AsiaPacific(Mumbai)\":901.45,\"Total\":1268.02,\"ServiceName\":\"AmazonElastiCache\",\"AsiaPacific(Singapore)\":16.9,\"AsiaPacific(Sydney)\":83.49,\"AsiaPacific(Tokyo)\":43.77,\"EU(Frankfurt)\":83.49,\"USEast(N.Virginia)\":138.92},{\"AsiaPacific(Mumbai)\":699.18,\"Total\":1622.45,\"ServiceName\":\"AmazonRelationalDatabaseService\",\"AsiaPacific(Singapore)\":103.59,\"AsiaPacific(Sydney)\":243.78,\"EU(Frankfurt)\":227.3,\"Miscellaneous\":43.15,\"USEast(N.Virginia)\":305.45},{\"AsiaPacific(Mumbai)\":272.26,\"Total\":374,\"ServiceName\":\"AWSLambda\",\"AsiaPacific(Singapore)\":3.64,\"AsiaPacific(Sydney)\":0.64,\"AsiaPacific(Tokyo)\":59.95,\"EU(Frankfurt)\":4.16,\"Miscellaneous\":3.01,\"USEast(N.Virginia)\":30.34},{\"AsiaPacific(Mumbai)\":101.15,\"Total\":511.32,\"ServiceName\":\"AmazonDynamoDB\",\"AsiaPacific(Singapore)\":6.06,\"AsiaPacific(Sydney)\":0.6,\"AsiaPacific(Tokyo)\":0.11,\"EU(Frankfurt)\":0.19,\"Miscellaneous\":399.32,\"USEast(N.Virginia)\":3.89},{\"AsiaPacific(Mumbai)\":33.8,\"Total\":202.8,\"ServiceName\":\"AmazonElasticContainerServiceforKubernetes\",\"AsiaPacific(Singapore)\":33.8,\"AsiaPacific(Sydney)\":33.8,\"AsiaPacific(Tokyo)\":33.8,\"EU(Frankfurt)\":33.8,\"USEast(N.Virginia)\":33.8},{\"AsiaPacific(Mumbai)\":16.06,\"Total\":71.37,\"ServiceName\":\"ElasticLoadBalancing\",\"AsiaPacific(Singapore)\":17.05,\"AsiaPacific(Sydney)\":8.52,\"AsiaPacific(Tokyo)\":12.87,\"EU(Frankfurt)\":9.14,\"Miscellaneous\":0.03,\"USEast(N.Virginia)\":7.7},{\"AsiaPacific(Mumbai)\":1.71,\"Total\":1.9,\"ServiceName\":\"AWSIoT\",\"AsiaPacific(Singapore)\":0.03,\"AsiaPacific(Sydney)\":0.01,\"AsiaPacific(Tokyo)\":0.09,\"EU(Frankfurt)\":0.03,\"USEast(N.Virginia)\":0.03},{\"AsiaPacific(Mumbai)\":1.45,\"Total\":11.86,\"ServiceName\":\"AmazonEC2ContainerRegistry(ECR)\",\"AsiaPacific(Singapore)\":2.36,\"AsiaPacific(Sydney)\":1.34,\"AsiaPacific(Tokyo)\":3.82,\"EU(Frankfurt)\":1.29,\"MiddleEast(Bahrain)\":0.23,\"Miscellaneous\":0.05,\"USEast(N.Virginia)\":1.32},{\"AsiaPacific(Mumbai)\":0.34,\"Total\":1.1,\"ServiceName\":\"AmazonAPIGateway\",\"AsiaPacific(Singapore)\":0.03,\"AsiaPacific(Sydney)\":0.01,\"AsiaPacific(Tokyo)\":0.16,\"EU(Frankfurt)\":0.04,\"Miscellaneous\":0,\"USEast(N.Virginia)\":0.52},{\"AsiaPacific(Mumbai)\":0.25,\"Total\":0.25,\"ServiceName\":\"AWSSecurityHub\"},{\"AsiaPacific(Mumbai)\":0,\"Total\":0,\"ServiceName\":\"AmazonCognito\"},{\"AsiaPacific(Mumbai)\":0,\"Total\":0,\"ServiceName\":\"AWSBackup\",\"AsiaPacific(Singapore)\":0,\"AsiaPacific(Sydney)\":0,\"EU(Frankfurt)\":0,\"USEast(N.Virginia)\":0},{\"AsiaPacific(Mumbai)\":0,\"Total\":0,\"ServiceName\":\"AWSGlue\",\"AsiaPacific(Tokyo)\":0,\"USEast(N.Virginia)\":0,\"USEast(Ohio)\":0},{\"AsiaPacific(Mumbai)\":0,\"Total\":0,\"ServiceName\":\"AWSSecretsManager\"},{\"AsiaPacific(Tokyo)\":7.14,\"Total\":8.14,\"ServiceName\":\"AWSConfig\",\"USEast(N.Virginia)\":1},{\"Miscellaneous\":0.2,\"Total\":0.2,\"ServiceName\":\"OpenVPNAccessServer(25ConnectedDevices)\"},{\"Miscellaneous\":0.19,\"Total\":0.19,\"ServiceName\":\"AmazonRoute53\"},{\"Miscellaneous\":0.03,\"Total\":7.26,\"ServiceName\":\"AmazonSimpleEmailService\",\"USWest(Oregon)\":7.23},{\"Miscellaneous\":0,\"Total\":0,\"ServiceName\":\"AmazonGlacier\"},{\"Miscellaneous\":0,\"Total\":0,\"ServiceName\":\"AWSElementalMediaStore\"},{\"USEast(N.Virginia)\":3.43,\"Total\":3.43,\"ServiceName\":\"AWSCostExplorer\"}],\"tableCurrencySymbolToShow\":[\"Africa(CapeTown)\",\"Any\",\"AsiaPacific\",\"AsiaPacific(HongKong)\",\"AsiaPacific(Mumbai)\",\"AsiaPacific(Osaka)\",\"AsiaPacific(Seoul)\",\"AsiaPacific(Singapore)\",\"AsiaPacific(Sydney)\",\"AsiaPacific(Tokyo)\",\"Australia\",\"Canada\",\"Canada(Central)\",\"EU(Frankfurt)\",\"EU(Ireland)\",\"EU(London)\",\"EU(Milan)\",\"EU(Paris)\",\"EU(Stockholm)\",\"Europe\",\"India\",\"Japan\",\"MiddleEast\",\"MiddleEast(Bahrain)\",\"Miscellaneous\",\"SouthAfrica\",\"SouthAmerica\",\"SouthAmerica(SaoPaulo)\",\"USEast(N.Virginia)\",\"USEast(Ohio)\",\"USWest(N.California)\",\"USWest(Oregon)\",\"UnitedStates\",\"Total\"],\"currencySymbol\":\"$\",\"preSignedURL\":\"https://s3.ap-south-1.amazonaws.com/centilytics.config.ap-south-1/temp-dir-lambda/centilytics_india_root/dev_cent%40centilytics.com/Regional%20Cost%20of%20AWS%20Services/Regional%20Cost%20of%20AWS%20Services.xlsx?X-Amz-Security-Token=IQoJb3JpZ2luX2VjEPH%2F%2F%2F%2F%2F%2F%2F%2F%2F%2FwEaCmFwLXNvdXRoLTEiSDBGAiEAuiGCjNIgKLs5%2BLvthoGDgOmhnj0xwe5kNhet5cDSfHoCIQDBxecCe1BnYzUwwt4Ok2Q4PnP1lnynCsAMvNcrLXGrpCquAgjq%2F%2F%2F%2F%2F%2F%2F%2F%2F%2F8BEAQaDDA2MDA4NzIxODE0NSIMJ62U5v5kSL4mNvF%2FKoICAe7dUD9pJBak7hx%2BmGx1znuCOQENABguxf%2FD%2FzVc%2BtSU9O6Ir3htBQDcAb6SL8TGhMJpUSWWMSlT9UU6W7eGGg8JwDiESuvnMXhel7FJUj24l97ERLPRw6Wz0dYtuSWzQc48WhO2YVb3iq3M9dtUHm0tAAfe2%2FcAI%2FCNL8ZD2l03zM%2FKWEMfupyY1Jr%2BLbh%2B4oJl%2F%2BjnQV%2FWSFgNTcv%2B1YVlQVRqOQPc%2FNRNUJ1x2FoN1cav0MucHbUI%2FrlndhwQT1E8hWsBP24QJrHLY2NahNUip5L3pqIQLB7%2BYtBAUV3KGKHe%2BMsDHTvPz1lz6OLdVK77EVnFR1FHXXLonmfiOF8FMPG%2FgZUGOpkBap%2F4FwGTjYTvql8MWQN9HQhaxNPfWeQBHwroLwMIDT%2ByuxFthF6%2FP9rOmR0T6u%2FaKvmoaCo62O2la1XniROfpjIieInXH3zQ%2FYXuHMzjvFoRo3KQyvVShTEhnwqw%2BP%2BywCd7FlNZABDfm02alSCL5AcrMPe0qJobeukuaG4%2B0cx9U2JKedtsgmRg72hgF3snZH1JYXIFg%2B%2Fq&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20220608T083838Z&X-Amz-SignedHeaders=host&X-Amz-Expires=86400&X-Amz-Credential=ASIAQ37L2F7QUJOQJPF6%2F20220608%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Signature=75b4f32373f8e6bb5ccb56e4e5ac671053b515569b3ecb6802007c28912ceb28\"},\"dataList\":[\"centilytics.config.ap-south-1/temp-dir-lambda/centilytics_india_root/dev_cent@centilytics.com/RegionalCostofAWSServices/RegionalCostofAWSServices.xlsx\"],\"noCredentialsDB\":[],\"insufficientPermission\":[],\"invalidCredentials\":[],\"generalException\":[],\"resourceTags\":[],\"self\":\"https://s3.ap-south-1.amazonaws.com/centilytics.config.ap-south-1/request-response-logs/response/centilytics_india_root/dev_cent%40centilytics.com/8d313f45-adfd-4649-b83b-d4d676106c21.json?X-Amz-Security-Token=IQoJb3JpZ2luX2VjEPH%2F%2F%2F%2F%2F%2F%2F%2F%2F%2FwEaCmFwLXNvdXRoLTEiSDBGAiEAuiGCjNIgKLs5%2BLvthoGDgOmhnj0xwe5kNhet5cDSfHoCIQDBxecCe1BnYzUwwt4Ok2Q4PnP1lnynCsAMvNcrLXGrpCquAgjq%2F%2F%2F%2F%2F%2F%2F%2F%2F%2F8BEAQaDDA2MDA4NzIxODE0NSIMJ62U5v5kSL4mNvF%2FKoICAe7dUD9pJBak7hx%2BmGx1znuCOQENABguxf%2FD%2FzVc%2BtSU9O6Ir3htBQDcAb6SL8TGhMJpUSWWMSlT9UU6W7eGGg8JwDiESuvnMXhel7FJUj24l97ERLPRw6Wz0dYtuSWzQc48WhO2YVb3iq3M9dtUHm0tAAfe2%2FcAI%2FCNL8ZD2l03zM%2FKWEMfupyY1Jr%2BLbh%2B4oJl%2F%2BjnQV%2FWSFgNTcv%2B1YVlQVRqOQPc%2FNRNUJ1x2FoN1cav0MucHbUI%2FrlndhwQT1E8hWsBP24QJrHLY2NahNUip5L3pqIQLB7%2BYtBAUV3KGKHe%2BMsDHTvPz1lz6OLdVK77EVnFR1FHXXLonmfiOF8FMPG%2FgZUGOpkBap%2F4FwGTjYTvql8MWQN9HQhaxNPfWeQBHwroLwMIDT%2ByuxFthF6%2FP9rOmR0T6u%2FaKvmoaCo62O2la1XniROfpjIieInXH3zQ%2FYXuHMzjvFoRo3KQyvVShTEhnwqw%2BP%2BywCd7FlNZABDfm02alSCL5AcrMPe0qJobeukuaG4%2B0cx9U2JKedtsgmRg72hgF3snZH1JYXIFg%2B%2Fq&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20220608T083838Z&X-Amz-SignedHeaders=host&X-Amz-Expires=172800&X-Amz-Credential=ASIAQ37L2F7QUJOQJPF6%2F20220608%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Signature=d82b6642e94f14eec984f06a1dd1ff441b4b5c76e50dab3dbe97e717dc1278d7\",\"next\":{\"baseUrl\":null,\"inputKeys\":null},\"previous\":{\"baseUrl\":null,\"inputKeys\":null},\"cloud\":\"aws\",\"invocationType\":null,\"module\":\"costmonitoring\",\"page\":\"home\",\"insight\":\"cost-by-region\",\"insightText\":\"RegionalCostofAWSServices\",\"referS3UrlBydefault\":false,\"basicLimit\":4,\"filterLimit\":null,\"insightState\":\"PREMIUM\",\"autoRenew\":false,\"otherException\":[],\"restrictedFilters\":[],\"message\":null}";

		
			try {

				JSONObject jsonObject = new JSONObject(dummyResponse);
				// System.out.println(jsonObject.getJSONObject("dataMap"));
				// System.out.println(jsonObject.getJSONObject("dataMap").get("table"));
				RunFunctionResponse response = jsonObject;


				// System.out.println(jsonObject);
				ObjectMapper mapper = new ObjectMapper();
				mapper.configure(SerializationFeature.FAIL_ON_EMPTY_BEANS, false);
				// LinkedList<String> keys2 = mapper.convertValue(jsonObject.getJSONObject("dataMap").get("table"),
				// new TypeReference<LinkedList<String>>() {
				// });
				// List<Map<String, Object>> table2 = mapper.convertValue(jsonObject.getJSONObject("dataMap").getJSONArray("table"),
				// 			new TypeReference<List<Map<String, Object>>>() {
				// 			});\\\
				// JSONArray dummy = jsonObject.getJSONObject("dataMap").getJSONArray("tableKeys");
				// List<String> table2 = dummy.toList();
				// System.out.println(table2);
				// HashMap<String,Object> response = mapper.convertValue(jsonObject,
				// 		new TypeReference<HashMap<String,Object>>() {});
				// System.out.println(response);

				// HashMap<String,Object> responseDataMap = mapper.convertValue(response.get("dataMap"),
				// 		new TypeReference<HashMap<String,Object>>() {});
				// System.out.println(responseDataMap);

		
			
				// FileWriter licenseFile = new FileWriter(Helper.getTempPath() + "Aspose.Cells.Product.Family.lic");
				// licenseFile.write(licenseData);
				// licenseFile.close();

				// FileInputStream fileIS = new FileInputStream(Helper.getTempPath() + "/Aspose.Cells.Product.Family.lic");
				// License license = new License();
				// license.setLicense(fileIS);

				SimpleDateFormat formatter = new SimpleDateFormat("dd MMM yyyy HH:mm:ss z");

				// timezoneFormat = DateTimeZone.forID(request.getUserInput().getTimezone());
				// Date startdate = (request.getUserInput().getStartDateTime() != null
				// 		&& !request.getUserInput().getStartDateTime().isEmpty())
				// 				? new SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
				// 						.parse(request.getUserInput().getStartDateTime())
				// 				: DateTime.now(timezoneFormat).minusDays(7).withTimeAtStartOfDay().toDate();
				// String startDate = new SimpleDateFormat("dd MMMM yyyy").format(startdate);
				// Date enddate = (request.getUserInput().getEndDateTime() != null
				// 		&& !request.getUserInput().getEndDateTime().isEmpty())
				// 				? new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(request.getUserInput().getEndDateTime())
				// 				: DateTime.now(timezoneFormat).withTimeAtStartOfDay().minusSeconds(1).toDate();
				// String endDate = new SimpleDateFormat("dd MMMM yyyy").format(enddate);

				// if (enddate.before(startdate)) {
				// 	throw new InputBasedException("End Date :- " + endDate + " Cannot Be Before Start Date" + startDate);
				// }


				Workbook workbook = new Workbook();
				Worksheet chartWorksheet = workbook.getWorksheets().get(0);
				int index = workbook.getWorksheets().add();
				Worksheet tableWorksheet = workbook.getWorksheets().get(index);

				// String tableSheetName = request.getUserInput().getReportSheetName();
				String tableSheetName = null;
				String chartSheetName = "Overview";
				// if (request.getUserInput().getReportSheetName() == null) {
				// 	tableSheetName = request.getApiConfig().getApiDefinition().getText();
				// }

				// if(response.containsKey("insightText")){
					tableSheetName = String.valueOf(response.insightText);

					// tableSheetName = String.valueOf(response.get("insightText"));
				// }	
				
				// if(jsonObject.getString("insightText")){
				// 	tableSheetName = String.valueOf(response.get("insightText"));
				// }		
				
				tableSheetName = CellsHelper.createSafeSheetName(tableSheetName, '_');
				tableSheetName = tableSheetName.trim();

				chartWorksheet.setName(chartSheetName);
				tableWorksheet.setName(tableSheetName);

				Cells cells = tableWorksheet.getCells();
				Cells chartCells = chartWorksheet.getCells();
				Style style = cells.get(0, 0).getStyle();
				Style defaultTopStyle = cells.get(1, 0).getStyle();
				style.setVerticalAlignment(TextAlignmentType.LEFT);
				style.setHorizontalAlignment(TextAlignmentType.LEFT);
				style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
				style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
				style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
				style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
				style.setForegroundColor(Color.fromArgb(242, 242, 242));
				style.getFont().setName("Calibri");
				style.getFont().setSize(10);
				style.setShrinkToFit(true);
				Style headingStyle = getStyleForHeading(cells.get(0, 0).getStyle());
				Font font = style.getFont();
				font.setBold(true);
				int row = 6;
				int column = 1;
				int maxColumn = 1;
				// System.out.println(response.get("dataMap"));
				if (response != null && !response.dataMap.isEmpty()) {
					LinkedList<String> keys = new ObjectMapper().convertValue(response.dataMap.get("keys"),
							new TypeReference<LinkedList<String>>() {
							});
					LinkedList<String> tableKeys = new ObjectMapper().convertValue(response.dataMap.get("tableKeys"),
							new TypeReference<LinkedList<String>>() {
							});
					List<Map<String, Object>> table = new ObjectMapper().convertValue(response.dataMap.get("table"),
							new TypeReference<List<Map<String, Object>>>() {
							});
					String currencySymbol = (String) response.dataMap.get("currencySymbol");
					String legend = tableKeys.getFirst();
					cells.merge(row, 0, 2, 1);
					cells.get(row, 0).getMergedRange().setValue("S.no.");
					cells.get(row, 0).getMergedRange().setStyle(headingStyle);
					cells.merge(row, 1, 2, 1);
					cells.get(row, 1).getMergedRange().setValue(legend);
					cells.get(row, 1).getMergedRange().setStyle(headingStyle);
					row++;
					String xAxisStart = cells.get(row, column + 1).getName();
					for (String tableKey : tableKeys) {
						if (tableKey.equals(legend))
							continue;
						column++;
						cells.get(row, column).setValue(tableKey);
						cells.get(row, column).setStyle(headingStyle);
					}
					String xAxisEnd = cells.get(row, column - 1).getName();
					row--;
					cells.merge(row, 2, 1, column - 1);
					cells.get(row, 2).getMergedRange().setValue("TimeStamp");
					cells.get(row, 2).getMergedRange().setStyle(headingStyle);
					row = row + 2;
					int count = 0;

					String legendStart = cells.get(row + 1, 1).getName();
					String dataRangeStart = cells.get(row + 1, 2).getName();
					for (String key : keys) {
						cells.get(row, 0).setValue(++count);
						cells.get(row, 0).setStyle(style);
						cells.get(row, 1).setValue(key);
						cells.get(row, 1).setStyle(style);
						for (int i = 2; i < column; i++) {
							cells.get(row, i).setValue("#N/A");
							cells.get(row, i).setStyle(style);
						}
						row++;
					}
					String legendEnd = cells.get(row - 1, 1).getName();
					String dataRangeEnd = cells.get(row - 1, column - 1).getName();
					double totalCost = 0;
					for (Map<String, Object> data : table) {
						if (((String) data.get(legend)).equalsIgnoreCase("Total"))
							totalCost += (double) data.get("Total");
						for (String oneRowData : data.keySet()) {
							int matrixColumn = 1;
							int matrixRow = row - count;
							if (!oneRowData.equalsIgnoreCase(legend)) {
								matrixColumn = matrixColumn + tableKeys.indexOf(oneRowData);
								matrixRow = matrixRow + keys.indexOf(data.get(legend));
								cells.get(matrixRow, matrixColumn).setValue(data.get(oneRowData));
								cells.get(matrixRow, matrixColumn).setStyle(style);
							}
						}
					}
					cells.merge(3, 0, 1, 2);
					cells.get(3, 0).getMergedRange().setValue("Total Cost");
					cells.get(3, 0).getMergedRange().setStyle(headingStyle);

					cells.merge(4, 0, 1, 2);
					cells.get(4, 0).getMergedRange().setValue(currencySymbol + totalCost);
					Style rightStyle = style;
					rightStyle.setVerticalAlignment(TextAlignmentType.RIGHT);
					rightStyle.setHorizontalAlignment(TextAlignmentType.RIGHT);
					cells.get(4, 0).getMergedRange().setStyle(rightStyle);

					tableWorksheet.autoFitColumns();
					tableWorksheet.autoFitRows();

					boolean nSeriesLimitFlag = table.size() > 255;

					if (nSeriesLimitFlag) {

						maxColumn = tableWorksheet.getCells().getMaxColumn() + 1;

						cells.merge(1, 0, 1, maxColumn);
						chartCells.merge(1, 0, 1, maxColumn);
						chartCells.merge(3, 0, 1, maxColumn);
						chartCells.get(3, 0).setValue("Graph is not supported for this range of data");
						chartCells.get(3, 0).setStyle(getStyleForNote(defaultTopStyle));

					} else {
						
						int chartIndex = chartWorksheet.getCharts().add(ChartType.COLUMN_STACKED, 4, 0, 30, 20);
						Chart chart = chartWorksheet.getCharts().get(chartIndex);
						int neriesIndex = chart.getNSeries().add("='" + tableWorksheet.getName() + "'!"
								+ convertNameToCellName(dataRangeStart) + ":" + convertNameToCellName(dataRangeEnd), false);
						int legendStartRange = Integer.parseInt(legendStart.substring(1));
						int legendEndRnage = Integer.parseInt(legendEnd.substring(1));
						int j = neriesIndex;
						for (int i = legendStartRange; i <= legendEndRnage; i++) {
							chart.getNSeries().get(j)
									.setName("='" + tableWorksheet.getName() + "'!$" + legendStart.substring(0, 1) + i);
							j++;
						}
						chart.getNSeries().setCategoryData("='" + tableWorksheet.getName() + "'!"
								+ convertNameToCellName(xAxisStart) + ":" + convertNameToCellName(xAxisEnd));
						chart.getNSeries().get(neriesIndex).setType(ChartType.COLUMN_STACKED);
						maxColumn = tableWorksheet.getCells().getMaxColumn() + 1;
						cells.merge(1, 0, 1, maxColumn);
						chartCells.merge(1, 0, 1, maxColumn);
						chartCells.merge(32, 0, 1, maxColumn);
						chartCells.get(32, 0).getMergedRange()
								.setValue("Note: To see all the legends for the graph, kindly extend the graph.");
						chartCells.get(32, 0).getMergedRange().setStyle(getStyleForNote(defaultTopStyle));
					}
					// if (request.getUserInput().getReportSheetName() == null
					// 		|| request.getUserInput().getReportSheetName().isEmpty()) {
					// 	cells.get(1, 0).getMergedRange().setValue(request.getApiConfig().getApiDefinition().getText() + " "
					// 			+ "( " + startDate + " - " + endDate + " )");
					// 	cells.get(1, 0).getMergedRange().setStyle(getStyleForTop(defaultTopStyle));
					// 	chartCells.get(1, 0).getMergedRange().setValue(request.getApiConfig().getApiDefinition().getText()
					// 			+ " " + "( " + startDate + " - " + endDate + " )");
					// 	chartCells.get(1, 0).getMergedRange().setStyle(headingStyle);
					// } else {
					// 	cells.get(1, 0).getMergedRange().setValue(request.getUserInput().getReportSheetName().substring(2)
					// 			+ " " + "( " + startDate + " - " + endDate + " )");
					// 	cells.get(1, 0).getMergedRange().setStyle(getStyleForTop(defaultTopStyle));
					// 	chartCells.get(1, 0).getMergedRange().setValue(request.getApiConfig().getApiDefinition().getText()
					// 			+ " " + "( " + startDate + " - " + endDate + " )");
					// 	chartCells.get(1, 0).getMergedRange().setStyle(headingStyle);
					// }

				} else {
					cells.get(3, 0).setValue("No Rows Available For The Given Set Of Inputs");
					cells.get(3, 0).setStyle(style);
					chartCells.get(3, 0).setValue("No Rows Available For The Given Set Of Inputs");
					chartCells.get(3, 0).setStyle(style);

					// if (request.getUserInput().getReportSheetName() == null
					// 		|| request.getUserInput().getReportSheetName().isEmpty()) {
					// 	cells.get(1, 0).setValue(request.getApiConfig().getApiDefinition().getText() + " " + "( "
					// 			+ startDate + " - " + endDate + " )");
					// 	cells.get(1, 0).setStyle(getStyleForTop(defaultTopStyle));
					// 	chartCells.get(1, 0).setValue(request.getApiConfig().getApiDefinition().getText() + " " + "( "
					// 			+ startDate + " - " + endDate + " )");
					// 	chartCells.get(1, 0).setStyle(headingStyle);
					// } else {
					// 	cells.get(1, 0).setValue(request.getUserInput().getReportSheetName().substring(2) + " " + "( "
					// 			+ startDate + " - " + endDate + " )");
					// 	cells.get(1, 0).setStyle(getStyleForTop(defaultTopStyle));
					// 	chartCells.get(1, 0).setValue(request.getApiConfig().getApiDefinition().getText() + " " + "( "
					// 			+ startDate + " - " + endDate + " )");
					// 	chartCells.get(1, 0).setStyle(headingStyle);
					// }
					tableWorksheet.autoFitColumns();
					chartWorksheet.autoFitColumns();
				}
				cells.get(++row, 0).setValue("Report Generated at: " + formatter.format(new Date()));
				// cells.get(++row, 0).setValue("Start Date: " + startDate + " " + request.getUserInput().getTimezone());
				// cells.get(++row, 0).setValue("End Date: " + endDate + " " + request.getUserInput().getTimezone());
				tableWorksheet.freezePanes(2, 0, 2, tableWorksheet.getCells().getMaxColumn() + 1);
				workbook.save(baos , SaveFormat.XLSX);
				buffer = baos.toByteArray();
				final String accessKey = "";
				final String secretKey = "";
		
				String bucketName = "centilytics-reporting";
				
				String fileName = "report2.xlsx";
		
				AwsBasicCredentials awsCreds = AwsBasicCredentials.create(accessKey,secretKey);
		
				S3Client client = S3Client.builder()
					.credentialsProvider(StaticCredentialsProvider.create(awsCreds))
					.build();
		
				PutObjectRequest request = PutObjectRequest.builder()
									.bucket(bucketName).key(fileName).build();
				
				client.putObject(request, RequestBody.fromBytes(buffer));
		
				// createFileInS3(tableWorksheet.getName());
			} catch (Exception e) {
				e.printStackTrace();
			}



			routingContext.response()
			.putHeader("content-type", "application/json;charset=UTF-8")
			.end(
			new JsonObject()
				.put("status", "ok")
				.encodePrettily()
			);
	}
  
	public Style getStyleForHeadings(Style style) {
		style.setVerticalAlignment(TextAlignmentType.CENTER_ACROSS);
		style.setHorizontalAlignment(TextAlignmentType.CENTER_ACROSS);

		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		Font font = style.getFont();
		font.setName("Calibri");
		font.setSize(11);
		font.setBold(true);
		style.getFont().setColor(Color.getWhite());
		style.setPattern(BackgroundType.SOLID);
		style.setShrinkToFit(true);
		return style;
	}

	public Style getStyleForCells(Style style) {
		style.setVerticalAlignment(TextAlignmentType.CENTER);
		style.setHorizontalAlignment(TextAlignmentType.CENTER);
		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.HAIR);
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.HAIR);
		Font font = style.getFont();
		font.setName("Calibri");
		font.setSize(11);
		font.setBold(false);
		font.setName("Calibri");
		style.setForegroundColor(Color.getWhite());
		style.setPattern(BackgroundType.SOLID);
		style.setShrinkToFit(true);
		return style;
	}

	public Style getStyleForTop(Style topStyle) {
		topStyle.setVerticalAlignment(TextAlignmentType.CENTER);
		topStyle.setHorizontalAlignment(TextAlignmentType.CENTER);
		topStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.setPattern(BackgroundType.SOLID);
		topStyle.getBorders().setColor(Color.fromArgb(150, 179, 215));
		topStyle.getFont().setName("Calibri");
		topStyle.getFont().setSize(13);
		topStyle.getFont().setBold(false);
		topStyle.getFont().setColor(Color.getBlack());
		topStyle.setShrinkToFit(true);
		topStyle.setForegroundColor(Color.fromArgb(255, 255, 255));

		return topStyle;
	}

	public Style getStyleForNote(Style topStyle) {
		topStyle.setVerticalAlignment(TextAlignmentType.LEFT);
		topStyle.setHorizontalAlignment(TextAlignmentType.LEFT);
		topStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		topStyle.setPattern(BackgroundType.SOLID);
		topStyle.getFont().setColor(Color.getWhite());
		topStyle.getFont().setSize(12);
		topStyle.getFont().setName("Calibri");
		topStyle.setShrinkToFit(true);
		topStyle.getFont().setBold(true);
		topStyle.setForegroundColor(Color.fromArgb(150, 179, 215));

		return topStyle;
	}

	public Style getStyleForTotalLeft(Style headingStyle) {
		headingStyle.setVerticalAlignment(TextAlignmentType.CENTER);
		headingStyle.setHorizontalAlignment(TextAlignmentType.LEFT);
		headingStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.setPattern(BackgroundType.SOLID);
		headingStyle.getBorders().setColor(Color.fromArgb(68, 202, 219));
		headingStyle.getFont().setColor(Color.getWhite());
		headingStyle.getFont().setSize(12);
		headingStyle.getFont().setName("Calibri");
		headingStyle.setShrinkToFit(true);
		headingStyle.getFont().setBold(true);
		headingStyle.setForegroundColor(Color.fromArgb(68, 202, 219));

		return headingStyle;
	}

	public Style getStyleForTotal(Style headingStyle) {
		headingStyle.setVerticalAlignment(TextAlignmentType.CENTER);
		headingStyle.setHorizontalAlignment(TextAlignmentType.CENTER);
		headingStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.setPattern(BackgroundType.SOLID);
		headingStyle.getBorders().setColor(Color.fromArgb(68, 202, 219));
		headingStyle.getFont().setColor(Color.getWhite());
		headingStyle.getFont().setSize(12);
		headingStyle.getFont().setName("Calibri");
		headingStyle.setShrinkToFit(true);
		headingStyle.getFont().setBold(true);
		headingStyle.setForegroundColor(Color.fromArgb(68, 202, 219));

		return headingStyle;
	}

	public Style getStyleForHeadingLeft(Style headingStyle) {
		headingStyle.setVerticalAlignment(TextAlignmentType.CENTER);
		headingStyle.setHorizontalAlignment(TextAlignmentType.LEFT);
		headingStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.setPattern(BackgroundType.SOLID);
		headingStyle.getBorders().setColor(Color.fromArgb(23, 55, 94));
		headingStyle.getFont().setColor(Color.getWhite());
		headingStyle.getFont().setSize(12);
		headingStyle.getFont().setName("Calibri");
		headingStyle.setShrinkToFit(true);
		headingStyle.getFont().setBold(true);
		headingStyle.setForegroundColor(Color.fromArgb(23, 55, 94));

		return headingStyle;
	}

	public Style getStyleForHeading(Style headingStyle) {
		headingStyle.setVerticalAlignment(TextAlignmentType.CENTER);
		headingStyle.setHorizontalAlignment(TextAlignmentType.CENTER);
		headingStyle.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		headingStyle.setPattern(BackgroundType.SOLID);
		headingStyle.getBorders().setColor(Color.fromArgb(23, 55, 94));
		headingStyle.getFont().setColor(Color.getWhite());
		headingStyle.getFont().setSize(12);
		headingStyle.getFont().setName("Calibri");
		headingStyle.setShrinkToFit(true);
		headingStyle.getFont().setBold(true);
		headingStyle.setForegroundColor(Color.fromArgb(23, 55, 94));

		return headingStyle;
	}

	public Style getStyleForOddCell(Style style) {
		style.setVerticalAlignment(TextAlignmentType.LEFT);
		style.setHorizontalAlignment(TextAlignmentType.LEFT);

		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.HAIR);
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.HAIR);
		style.getBorders().setColor(Color.fromArgb(150, 179, 215));
		Font font = style.getFont();
		font.setSize(10);
		font.setBold(false);
		style.setForegroundColor(Color.fromArgb(242, 242, 242));
		style.setPattern(BackgroundType.SOLID);
		style.setShrinkToFit(true);
		return style;
	}

	public Style getStyleForEvenCell(Style style) {
		style.setVerticalAlignment(TextAlignmentType.LEFT);
		style.setHorizontalAlignment(TextAlignmentType.LEFT);

		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.HAIR);
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.HAIR);
		style.getBorders().setColor(Color.fromArgb(150, 179, 215));
		style.setForegroundColor(Color.fromArgb(242, 242, 242));
		Font font = style.getFont();
		font.setSize(10);
		font.setBold(false);
		style.setForegroundColor(Color.getWhite());
		style.setPattern(BackgroundType.SOLID);
		style.setShrinkToFit(true);
		return style;
	}

	private Style getCommonStyle(Style style) {
        style.setVerticalAlignment(TextAlignmentType.LEFT);
        style.setHorizontalAlignment(TextAlignmentType.LEFT);
        style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
        style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
        style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
        style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
        style.getFont().setName("Calibri");
        style.getFont().setSize(10);
        style.setShrinkToFit(true);
        return style;
    }

	private String convertNameToCellName(String s) {
		String[] part = s.split("(?<=\\D)(?=\\d)");
		return "$" + part[0] + "$" + part[1];
	}

}
