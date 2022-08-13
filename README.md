# VBA_Challenge
 ## Overview
 The purpose of this project is to refactor a script within the worst language known to mankind, Visual Basic, that will output gains or losses in stock values over time at a faster rate than the ~half a second the original code takes to run on my computer.
 If this project weren't shorter than my other missing assignments I would have left it as one of my missing assignments. If I never open the Visual Basic editor again I will consider myself as having lived a happy life. I fought the compiler and the compiler won.
 ## Results
 The results of this project are that the original code ran in under half a second for both years. My computer is speedy I guess.  
 **2017 original code results**  
 ![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/OldCode2017.png)  
 **2018 original code results**
 ![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/OldCode2018.png)  
The results for the refactored code is an overflow error, which resulted from trying to fix a type mismatch error, which I believe resulted from trying to fix an error about being unable to assign a value to the stock index array, which resulted from the compiler telling me it expected stock index to be an array in step 3a.  
**Refactored code results (both years)**  
![](https://github.com/ChrisJAnderson/VBA_Challenge/blob/main/Resources/RefactoredCodeError.png)  
## Summary
Please don't take my unwillingness to continue troubleshooting this code as a lack of understanding about the usefulness of refactoring code. Being able to fine-tune or retailor code for one's own purposes is generally  highly useful. It's much quicker and less work intensive to refactor existing code to accomodate for a larger scale of information, a slightly different output, or to fine-tune it to get a quicker runtime than writing a new program from scratch. I see it as akin to using a form letter to address a client versus writing a new email every time- the skeleton is there, you know it works and that there are no outstanding problems. Therefore you can devote your time to addressing the things that are important- in the case of a form letter the client's name, title, place of work, etc and in the case of code the scale and runtime of the script without reinventing the wheel each time.
The drawbacks to refactoring code should be obvious in the scope of this project.
Sometimes a program is "good enough" and doesn't need to be tweaked or fine-tuned, or shouldn't be refactored by someone who doesn't know what they're doing. In my case, for example, I spent six hours trying to shave fractions of a second off of a macro that was perfectly serviceable, just a little slow. In that six hours, that macro could have run more than 43 thousand times given its runtime on my pc.  
  
I cannot honestly give a detailed statement on the advantages or disadvantages of the refactored script versus the original, as my refactored script does not run.
I suppose the advantage of the original over the refactor is that the original works.