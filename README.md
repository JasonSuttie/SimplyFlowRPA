# SimplyFlowRPA
RPA for Windows 10+

Welcome to Simply Flow RPA for Windows 10 or higher. 

Learn how to write RPA automation scripts using this library and get automating in 20 minutes or less. This library is aimed at developers who need RPA as part of their overall solution. It's simple to understand and use, people with less experience can learn how to use it too. This library has been used on multiple RPA solutions worldwide and is used in 24/7 scenarios.

Here are some tips to getting the best results:
- Develop and run your RPA scripts at the same screen resolution, lower color depth is better
- Some legacy Windows apps may require that they are launched with administrator rights, bear that in mind
- The library uses Windows OCR as a simple form of computer vision. Keep your automation sequences short, run them and then "look" for the expected result
- RPA scripts mean that the user interface must be available, virtual machines work well and can be better secured
- Play with how you capture desktop images by converting them to black and white to get the best OCR result. Use "ReturnWordLocationAdvanced"
- A good pattern to implement is web service + Windows task scheduler + RPA script to allow input data to be pushed using the web service and run on a schedule using task scheduler. Async responses can then be sent to the calling system.
- Yes, we've done this plenty and can help you



