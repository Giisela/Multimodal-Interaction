# Multimodal-Interaction
The project consists of using the power point from voice commands and gestures.

# Speech
This folder contains the entire process used to carry out voice commands. We use a drive that is found in the references of the visual studio that corresponds to the Power Point. We created the dynamic grammar in Portuguese (it was one of the requested requirements) found in speechMod and in appgui you can find all the programmed execution of the features. The project consists of two parts, presentation mode and editing mode, this means that it is possible to use features in both presentation mode and editing mode.
In editing mode we have:
* Open and close power point
* forward and backward slides
* jump to different slides
* change of theme
* add and remove new slide
* change title and text color
* read title, text or notes
* save changes

In presentation mode we have:
* forward and backward slides
* jump to different slides
* change of theme
* read title, text or notes
* open and close presentation mode

# Gestures
To make the gestures we use the programs available from kinect. We recorded the gestures and made the program learn with the aid of machine learning.
With gestures it is possible:
* forward and backward slide
* change theme
* open and close presentation mode
* zoom in / out
* crop in / out

# Fusion
As the name implies is the fusion between speech and gestures. In this we have some speech features and all gestures. We have simple, redundant and complementary mergers.

All projects have voice and gesture feedback.
