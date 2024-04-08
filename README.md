## Sample Application to verify fetching colors via UISettings' instance
The application contains basic implementation of getting the instance of `UISettings` class using `RoActivateInstance` and appropriate p/invoke wrappers.
The instance would be fetched on the button's click, and the color values can be verified by having adding a breakpoint at `Button_Click` event's implementation.
