




		window.addEventListener("load", Initiate, false);

		var Debugger = function () { };
		Debugger.log = function (message) {
		   try {
		      console.log(message);
		   } catch (exception) {
		      return;
		   }
		}

		function Initiate () {
		   Matrix();
		}


		// The code on every website is what glues the design together. For clarification, code is the markup and the programming languages that create the website