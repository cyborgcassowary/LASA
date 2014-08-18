var $ =jQuery.noConflict();
$(document).ready(function () {
  $('[data-toggle="offcanvas"]').click(function () {
    $('.row-offcanvas').toggleClass('active')
  });
});

var protocol = location.protocol;
	var host = location.host;
	var pathName = location.pathname;
	var lasaurl = host + pathName;


	function QueryFactory(DS, formName, Q){
		var DomSel = document.querySelector(DS).innerHTML; 
		var fE = [];

		var formElement = document.createElement('input');
			formElement.type = 'hidden';
		var FName = formElement.name = formName;
		var trimVar = formElement.value = encodeURIComponent(DomSel.trim());
		var theQuery = Q + trimVar;

	// Create an array to return the name of the hidden form element and part of the query string the the url.
		var D = [formElement, FName, theQuery];
			return{
				D: D
			};
	}

	function createForm(){
		// Create Form
		var FORM = document.createElement('div');
		FORM.id = 'requestcontainer';
		
		// button
		var BTN = document.createElement('a');
		var BTNtext=document.createTextNode("Request for Training");
		BTN.appendChild(BTNtext);
		//BTN.value = "Request for Training";

		var QueryString = protocol +'//'+ lasaurl + '?q=training/request-training'; 

		// Setting up Global Array variables for each possible field.
		var Employ33 = [];
		var JobTitl3 = [];
		var D3partment = [];
		var CourseNam3 = [];
		var Cat3 = [];
		var Cours3Sponsor = [];
		var Attendanc3Hours = [];
		var ContactHour3 = [];
		var R3gistration = [];
		var Lodging3 = [];
		var Trav3l = [];
		var M3als = [];
		var D3scription = [];
		var Oth3r = [];
		var R3ason = [];
		var Location3 = [];


		// Employee - 

		if(document.querySelector('.field-name-field-employee-3 .field-items .field-item a')){	
			Employ33.push(QueryFactory('.field-name-field-employee-3 .field-items .field-item a', 'employee', '&employee='));
			QueryString += Employ33[0].D[2] ; 
		} 
		if(document.querySelector('.field-name-field-employee-3 .links .taxonomy-term-reference-0 a')){
			Employ33.push(QueryFactory('.field-name-field-employee-3 .links .taxonomy-term-reference-0 a', 'employee', '&employee='));
			QueryString += Employ33[0].D[2] ; 
		}
		

		// Job title - field-name-field-job-title-training

		if(document.querySelector('.field-name-field-job-title-training .field-items .field-item')){	
			JobTitl3.push(QueryFactory('.field-name-field-job-title-training .field-items .field-item', 'jobtitle', '&jobtitle='));
			QueryString += JobTitl3[0].D[2] ; 
			// Fieldset.appendChild(JobTitl3[0].D[);
		} 
		if(document.querySelector('.field-name-field-job-title-training .links .taxonomy-term-reference-0')){
			JobTitl3.push(QueryFactory('.field-name-field-job-title-training .links .taxonomy-term-reference-0', 'jobtitle', '&jobtitle='));
			QueryString += JobTitl3[0].D[2] ; 
			// Fieldset.appendChild(JobTitl3[0].D[0] );
		}
		 
		// // Department -  field-name-field-department

		if(document.querySelector('.field-name-field-department .links .taxonomy-term-reference-0 a')){
			D3partment.push(QueryFactory('.field-name-field-department .links .taxonomy-term-reference-0 a', 'department', '&department='));
			QueryString += D3partment[0].D[2] ; 
			// Fieldset.appendChild(D3partment[0].D[0]);
		} 
		if(document.querySelector('.field-name-field-department .field-items .field-item a')){	
			D3partment.push(QueryFactory('.field-name-field-department .field-items .field-item a', 'department', '&department='));
			QueryString += D3partment[0].D[2] ; 
			// Fieldset.appendChild(D3partment[0].D[0]);
		}
		
		// //  else { 
		// // 	return null;
		// // }

		
		// // Course Name - field-name-field-coursename

		if(document.querySelector('.field-name-field-course-name .field-items .field-item a')){	
			CourseNam3.push(QueryFactory('.field-name-field-course-name .field-items .field-item a', 'coursename', '&coursename='));
			QueryString += CourseNam3[0].D[2] ; 
			// Fieldset.appendChild(CourseNam3[0].D[0]);
		} 	
		if(document.querySelector('.field-name-field-course-name .links .taxonomy-term-reference-0 a ')){
			CourseNam3.push(QueryFactory('.field-name-field-course-name .links .taxonomy-term-reference-0 a', 'coursename', '&coursename='));
			QueryString += CourseNam3[0].D[2] ; 
		}
		

		// // Training Category - field-name-field-training-category

		if(document.querySelector('.field-name-field-training-category .field-items .field-item a')){	
			Cat3.push(QueryFactory('.field-name-field-training-category .field-items .field-item a', 'category', '&category='));
			QueryString += Cat3[0].D[2] ; 
			// Fieldset.appendChild(Cat3[0].D[0]);
		} 	
		if(document.querySelector('.field-name-field-training-category .links .taxonomy-term-reference-0 a ')){
			Cat3.push(QueryFactory('.field-name-field-training-category .links .taxonomy-term-reference-0 a', 'category', '&category='));
			QueryString += Cat3[0].D[2] ; 
			// Fieldset.appendChild(Cat3[0].D[0]);
		}
		 

		// // // // Course Sponsor - field-name-field-course-sponsor

		if(document.querySelector('.field-name-field-course-sponser .field-items .field-item a')){	
			Cours3Sponsor.push(QueryFactory('.field-name-field-course-sponser .field-items .field-item a', 'sponsor', '&sponsor='));
			QueryString += Cours3Sponsor[0].D[2] ; 
		} 
		if(document.querySelector('.field-name-field-course-sponser .links .taxonomy-term-reference-0 a ')){
			Cours3Sponsor.push(QueryFactory('.field-name-field-course-sponser .links .taxonomy-term-reference-0 a', 'sponsor', '&sponsor='));
			QueryString += Cours3Sponsor[0].D[2] ; 
		}

		// // // // Attendance Hours - field-name-field-attendance-hours

		if(document.querySelector('.field-name-field-attendance-hours .field-items .field-item a')){	
			Attendanc3Hours.push(QueryFactory('.field-name-field-attendance-hours .field-items .field-item a', 'ahours', '&ahours='));
			QueryString += Attendanc3Hours[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-attendance-hours .links .taxonomy-term-reference-0 a ')){
			Attendanc3Hours.push(QueryFactory('.field-name-field-attendance-hours .links .taxonomy-term-reference-0 a', 'ahours', '&ahours='));
			QueryString += Attendanc3Hours[0].D[2] ; 
		}

		// // // // Contact Hours Amount - field-name-field-contact-hours-amount

		if(document.querySelector('.field-name-field-contact-hours-amount .field-items .field-item a')){	
			ContactHour3.push(QueryFactory('.field-name-field-contact-hours-amount .field-items .field-item a', 'contacthours', '&contacthours='));
			QueryString += ContactHour3[0].D[2] ; 
		}
		if(document.querySelector('.field-name-field-contact-hours-amount .links .taxonomy-term-reference-0 a ')){
			ContactHour3.push(QueryFactory('.field-name-field-contact-hours-amount .links .taxonomy-term-reference-0 a', 'contacthours', '&contacthours='));
			QueryString += ContactHour3[0].D[2] ; 
		}

		// // // // Registration field-name-field-registration

		if(document.querySelector('.field-name-field-registration .field-items .field-item a')){	
			R3gistration.push(QueryFactory('.field-name-field-registration .field-items .field-item a', 'registration', '&registration='));
			QueryString += R3gistration[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-registration .links .taxonomy-term-reference-0 a ')){
			R3gistration.push(QueryFactory('.field-name-field-registration .links .taxonomy-term-reference-0 a', 'registration', '&registration='));
			QueryString += R3gistration[0].D[2] ; 
		}
		 

		// // // Lodging field-name-field-lodging

		if(document.querySelector('.field-name-field-lodging .field-items .field-item a')){	
			Lodging3.push(QueryFactory('.field-name-field-lodging .field-items .field-item a', 'lodge', '&lodge='));
			QueryString += Lodging3[0].D[2] ; 
		} 
		if(document.querySelector('.field-name-field-lodging .links .taxonomy-term-reference-0 a ')){
			Lodging3.push(QueryFactory('.field-name-field-lodging .links .taxonomy-term-reference-0 a', 'lodge', '&lodge='));
			QueryString += Lodging3[0].D[2] ; 
		}
		 

		// // // // Travel field-name-field-travel

		if(document.querySelector('.field-name-field-travel .field-items .field-item a')){	
			Trav3l.push(QueryFactory('.field-name-field-travel .field-items .field-item a', 'travel', '&travel='));
			QueryString += Trav3l[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-travel .links .taxonomy-term-reference-0 a ')){
			Trav3l.push(QueryFactory('.field-name-field-travel .links .taxonomy-term-reference-0 a', 'travel', '&travel='));
			QueryString += Trav3l[0].D[2] ; 
		}
		 

		// // // Meals field-name-field-meals

		if(document.querySelector('.field-name-field-meals .field-items .field-item a')){	
			M3als.push(QueryFactory('.field-name-field-meals .field-items .field-item a', 'meals', '&meals='));
			QueryString += M3als[0].D[2] ; 
			// Fieldset.appendChild();
		} 	
		if(document.querySelector('.field-name-field-meals .links .taxonomy-term-reference-0 a ')){
			M3als.push(QueryFactory('.field-name-field-meals .links .taxonomy-term-reference-0 a', 'meals', '&meals='));
			QueryString += M3als[0].D[2] ; 
			// Fieldset.appendChild(M3als[0].D[0]);
		}  

		// // // Course Description field-name-field-course-description

		if(document.querySelector('.field-name-field-course-description .field-items .field-item a')){	
			D3scription.push(QueryFactory('.field-name-field-course-description .field-items .field-item a', 'cDescription', '&cDescription='));
			QueryString += D3scription[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-course-description .links .taxonomy-term-reference-0 a')){
			D3scription.push(QueryFactory('.field-name-field-course-description .links .taxonomy-term-reference-0 a', 'cDescription', '&cDescription='));
			QueryString += D3scription[0].D[2] ; 
		}
 

		// // // Other field-name-field-other

		if(document.querySelector('.field-name-field-other .field-items .field-item a')){	
			Oth3r.push(QueryFactory('.field-name-field-other .field-items .field-item a', 'otherexps', '&otherexps='));
			QueryString += Oth3r[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-other .links .taxonomy-term-reference-0 a ')){
			Oth3r.push(QueryFactory('.field-name-field-other .links .taxonomy-term-reference-0 a', 'otherexps', '&otherexps='));
			QueryString += Oth3r[0].D[2] ; 
		}
 

		// // // // Reason field-name-field-reason

		if(document.querySelector('.field-name-field-reason .field-items .field-item a')){	
			R3ason.push(QueryFactory('.field-name-field-reason .field-items .field-item a', 'reason', '&reason='));
			QueryString += R3ason[0].D[2] ; 
		} 	
		if(document.querySelector('.field-name-field-reason .links .taxonomy-term-reference-0 a ')){
			R3ason.push(QueryFactory('.field-name-field-reason .links .taxonomy-term-reference-0 a', 'reason', '&reason='));
			QueryString += R3ason[0].D[2] ; 
		}
 

		// // // Location Site field-name-field-location-site

		if(document.querySelector('.field-name-field-location-site .links .taxonomy-term-reference-0 a')){
			Location3.push(QueryFactory('.field-name-field-location-site .links .taxonomy-term-reference-0 a', 'location', '&location='));
			QueryString += Location3[0].D[2]  
		} 	
		if(document.querySelector('.field-name-field-location-site .field-items .field-item a')){	
			Location3.push(QueryFactory('.field-name-field-location-site .field-items .field-item a', 'location', '&location='));
			QueryString += Location3[0].D[2] ; 
		}

		//+'?q=training/request-training&'+'&jobtitle='+jobTVal+'&department='+Department+'&coursename='+CourseName+'&category='+Category+'&sponsor='+Sponsor+'&location='+Loc+'&coursedate='+CourseDate+'&ahours='+AHours+'&contacthours='+ContactHours+'&cDescription='+cDes+'&reason='+Reason+'&travel='+Travel+'&lodge='+Lodge+'&meals='+Meals+'&otherexps='+OtherExps;

var a = document.createAttribute("href");
  a.value = QueryString;
  BTN.setAttributeNode(a);
  var BTNlink = document.createAttribute("class");
  BTNlink.value = "btn btn-primary";
  BTN.setAttributeNode(BTNlink);
  //a.appendChild(BTN);

	
if(document.querySelector('.node-courses-and-conferences')){ FORM.appendChild(BTN); document.body.querySelector('.node-courses-and-conferences').appendChild(FORM); }
}

window.addEventListener("load", eventWindowLoaded, false);
function eventWindowLoaded () {
   createForm();
}