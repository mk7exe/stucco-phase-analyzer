# Stucco Phase Analyzer

The purpose of this project is to develope an accurate and reliable method to calculate phase contnet of commercial stucco. In drywall production, [stucco](https://en.wikipedia.org/wiki/Stucco) is refered to the raw material in form of a white powder which is the input to the production plant and is mainly made of calium sulfate hemihydrate (CaSO<sub>4</sub>.0.5H<sub>2</sub>O). In addition to calium sulfate hemihydrate, stucco may contain calium sulfate dihydrate (CaSO<sub>4</sub>.2H<sub>2</sub>O), calium sulfate anhydrate (CaSO<sub>4</sub>) and inert phases. In this context, inert phases are those that are not hydratable (can not adsorb water in their crystal structure). The quality of the final drywall product is highly depended on composition and phase content of the initial stucco. This project is divided into two phases:

* __Phase 1__: Developing a software to calculate phase contents through solving a system of linear equations. The linear equations are made from weight loss and weight gain data of three samples. The three samples are coded as follows:
  * ORG: The original stucco sample
  * HUM: The stucco sample conditioned at 75% humidity over night and dried for at least 2 hours at 45 C.
  * HYD: The stucco sample fully hydrated and then dried overnight at 45 C in a vaccum oven.
  
  The weight gain percentages of HUM and HYD samples are simply measured by weighing samples before and after conditioning using a high-precision scale. The weight loss percentages of each of the three sample is the measure of crystalline water which is removed at high temperature. This is usually measured by the TGA machine. 

* __Phase 2__: Developing a model based on neural networks that gives phase contents using two curves measured fot the stucco sample:
  * Temperature change during the hydration
  * Weight loss during the calcination
