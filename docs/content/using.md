## Using the application.

As previously stated the goal of the project is to enable aid project data reporting to the [IATI standard](http://iatistandard.org/).  The project originated from a requirement at WaterAid but it is likely to
be applicable for other organisations that need to meet the IATI standard.  The documentation in these pages describes the use of the application in the canonical form.  That may not be applicable for your 
organisation and consequently we recommend you contact [Mike Smith](https://github.com/drmrsmith) at [WaterAid](http://www.wateraid.org/uk/) to discuss if/ how you can change the application for your needs.
Mike can offer guidance but if you feel you don't need it please feel free to clone or fork the repo and use it in accordance with the licence terms.  We ask that if you do use it could you please STAR the repo so we can get 
an idea of the spread of the application.

###IATI Standard

At its base level the IATI Standard is enforced by an XML Schema.  Like all XML documents, a report that is valid according to the IATI standard is simply well-formed and valid against the XML schema.  
In the case of the IATI schema however the document can be very parsimonious or it can be very verbose, the schema is very permissive.  Validation against the schema is really no reassurance
that your reporting is 'correct'.  An additional level of 'validation' is provided by the IATI RuleSets which are JSON documents containing checks like: activity-enddate must be after activity-startdate.
It is via the rulesets that additional report validation will be enforced.  For more information on this subject have a look at the [validator][http://validator.iatistandard.org/].

####Validation

This tool guarantees schema validity but it does not guarantee complete ruleset validity.  The following is a list of validation the tool ensures:

* Schema v1 