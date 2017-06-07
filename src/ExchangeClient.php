<?php

use \jamesiarmes\PhpEws\Client;

use \jamesiarmes\PhpEws\Request\CreateItemType;
use \jamesiarmes\PhpEws\Request\FindItemType;
use \jamesiarmes\PhpEws\Request\UpdateItemType;

use \jamesiarmes\PhpEws\Response\FindItemResponseMessageType;

use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;
use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfBaseFolderIdsType;
use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfBaseItemIdsType;
use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfItemChangeDescriptionsType;
use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfPathsToElementType;

use \jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use \jamesiarmes\PhpEws\Enumeration\CalendarItemCreateOrDeleteOperationType;
use \jamesiarmes\PhpEws\Enumeration\CalendarItemUpdateOperationType;
use \jamesiarmes\PhpEws\Enumeration\ConflictResolutionType;
use \jamesiarmes\PhpEws\Enumeration\DefaultShapeNamesType;
use \jamesiarmes\PhpEws\Enumeration\DistinguishedFolderIdNameType;
use \jamesiarmes\PhpEws\Enumeration\DistinguishedPropertySetType;
use \jamesiarmes\PhpEws\Enumeration\ItemQueryTraversalType;
use \jamesiarmes\PhpEws\Enumeration\MapiPropertyTypeType;
use \jamesiarmes\PhpEws\Enumeration\MessageDispositionType;
use \jamesiarmes\PhpEws\Enumeration\ResponseClassType;
use \jamesiarmes\PhpEws\Enumeration\RoutingType;
use \jamesiarmes\PhpEws\Enumeration\UnindexedFieldURIType;

use \jamesiarmes\PhpEws\Type\BodyType;
use \jamesiarmes\PhpEws\Type\CalendarItemType;
use \jamesiarmes\PhpEws\Type\CalendarViewType;
use \jamesiarmes\PhpEws\Type\CancelCalendarItemType;
use \jamesiarmes\PhpEws\Type\DistinguishedFolderIdType;
use \jamesiarmes\PhpEws\Type\ExtendedPropertyType;
use \jamesiarmes\PhpEws\Type\FindItemParentType;
use \jamesiarmes\PhpEws\Type\ItemChangeType;
use \jamesiarmes\PhpEws\Type\ItemIdType;
use \jamesiarmes\PhpEws\Type\ItemResponseShapeType;
use \jamesiarmes\PhpEws\Type\PathToExtendedFieldType;
use \jamesiarmes\PhpEws\Type\PathToUnindexedFieldType;
use \jamesiarmes\PhpEws\Type\SetItemFieldType;


class ExchangeClient
{
    private $client = null;

    public function __construct($server, $username, $password)
    {
        $this->client = new Client($server, $username, $password, Client::VERSION_2016);
    }

    /**
     * @param int $daysFromNow
     * @return array
     */
    public function getCalendarEvents($daysFromNow)
    {
        // Set init class
        $request = new FindItemType();

        // Use this to search only the items in the parent directory in question or use ::SOFT_DELETED
        // to identify "soft deleted" items, i.e. not visible and not in the trash can.
        $request->Traversal = ItemQueryTraversalType::SHALLOW;

        // This identifies the set of properties to return in an item or folder response
        $request->ItemShape = new ItemResponseShapeType();
        $request->ItemShape->BaseShape = DefaultShapeNamesType::ALL_PROPERTIES;

        // Define the timeframe to load calendar items
        $request->CalendarView = new CalendarViewType();
        $request->CalendarView->StartDate = date(DATE_RFC3339);
        $request->CalendarView->EndDate = date(DATE_RFC3339, strtotime('+' . $daysFromNow . ' days'));

        // Only look in the "calendars folder"
        $request->ParentFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
        $request->ParentFolderIds->DistinguishedFolderId = new DistinguishedFolderIdType();
        $request->ParentFolderIds->DistinguishedFolderId->Id = DistinguishedFolderIdNameType::CALENDAR;

	// We want to get the Google Calendar ID in the response. Note that if this
	// property is not set on the event, it will not be included in the response.
	$property = new PathToExtendedFieldType();
	$property->PropertyName = 'GoogleCalendarID';
	$property->PropertyType = MapiPropertyTypeType::STRING;
	$property->DistinguishedPropertySetId = 'Appointment';
	$additional_properties = new NonEmptyArrayOfPathsToElementType();
	$additional_properties->ExtendedFieldURI[] = $property;
	$request->ItemShape->AdditionalProperties = $additional_properties;

        // Send request
        $response = $this->client->FindItem($request);

        // Loop through each item if event(s) were found in the timeframe specified
        $items = [];
        if ($response->ResponseMessages->FindItemResponseMessage[0]->RootFolder->TotalItemsInView > 0) {
            $events = $response->ResponseMessages->FindItemResponseMessage[0]->RootFolder->Items->CalendarItem;
	    foreach ($events as $event) {
                $items[$event->ItemId->Id] = [
                    'id' => $event->ItemId->Id,
                    'changeKey' => $event->ItemId->ChangeKey,
		    'googlecalendarid' => !empty($event->ExtendedProperty) ? $event->ExtendedProperty[0]->Value : null,
                    'subject' => $event->Subject,
                    'start' => $event->Start,
                    'end' => $event->End,
                    'location' => !empty($event->Location) ? $event->Location : null,
                    'isAllDayEvent' => !empty($event->IsAllDayEvent) ? true : false,
                    'isPublic' => strtolower($event->Sensitivity) == 'normal' ? true : false,
                    'isBusyStatus' => strtolower($event->LegacyFreeBusyStatus) == 'busy' ? true : false,
                    'isReminderSet' => !empty($event->ReminderIsSet) ? true : false,
                    'reminderMinutesBeforeStart' => !empty($event->ReminderMinutesBeforeStart) ? $event->ReminderMinutesBeforeStart : null
                ];
            }
        }
        return $items;
    }

    public function addEvent(array $details = []) {
	// Build the request,
        $request = new CreateItemType();
        $request->SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType::SEND_ONLY_TO_ALL;
        $request->Items = new NonEmptyArrayOfAllItemsType();

        // Build the event to be added.
        $event = new CalendarItemType();
	$start = new DateTime($details['start']);
	$end = new DateTime($details['end']);

        $event->Start = $start->format('c');
        $event->End = $end->format('c');
        $event->Subject = $details['subject'];
	$event->Location = $details['location'];

       // Set the event body.
//       $event->Body = new BodyType();
//       $event->Body->_ = $details[''];
//       $event->Body->BodyType = BodyTypeType::TEXT;

       // Add the GoogleCalendarID to extended properties
       $property = new ExtendedPropertyType();
       $property->ExtendedFieldURI = new PathToExtendedFieldType();
       $property->ExtendedFieldURI->DistinguishedPropertySetId = "Appointment";
       $property->ExtendedFieldURI->PropertyName = 'GoogleCalendarID';
       $property->ExtendedFieldURI->PropertyType = MapiPropertyTypeType::STRING;
       $property->Value = $details['googlecalendarid'];

       $event->ExtendedProperty = $property;

       // Add the event to the request. You could add multiple events to create more
       // than one in a single request.
       $request->Items->CalendarItem[] = $event;
       $response = $this->client->CreateItem($request);

       // Iterate over the results, printing any error messages or event ids.
       $response_messages = $response->ResponseMessages->CreateItemResponseMessage;
       foreach ($response_messages as $response_message) {
       // Make sure the request succeeded.
           if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
               $code = $response_message->ResponseCode;
               $message = $response_message->MessageText;
               fwrite(STDERR, "Event failed to create with \"$code: $message\"\n");
               continue;
           }
           // Iterate over the created events, printing the id for each.
           foreach ($response_message->Items->CalendarItem as $item) {
                $id = $item->ItemId->Id;
           }
        }
    }

    public function updateEvent($eventID, array $details = []) {
	// Build the request.
	$request = new UpdateItemType();
	$request->ConflictResolution = ConflictResolutionType::ALWAYS_OVERWRITE;
	$request->SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType::SEND_TO_NONE;

	// Build out item change request.
	$change = new ItemChangeType();
	$change->ItemId = new ItemIdType();
	$change->ItemId->Id = $eventID;
	$change->Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

	// Start field
	$start = new DateTime($details['start']);
	$field = $this->getChangeItemField('Start', $start->format('c'), UnindexedFieldURIType::CALENDAR_START);
	$change->Updates->SetItemField[] = $field;

	// End field
	$end = new DateTime($details['end']);
	$field = $this->getChangeItemField('End', $end->format('c'), UnindexedFieldURIType::CALENDAR_END);
	$change->Updates->SetItemField[] = $field;

	// Subject field
	$field = $this->getChangeItemField('Subject', $details['subject'], UnindexedFieldURIType::ITEM_SUBJECT);
	$change->Updates->SetItemField[] = $field;

	// Location field
	if ( $details['location'] ) {
		$field = $this->getChangeItemField('Location', $details['location'], UnindexedFieldURIType::CALENDAR_LOCATION);
		$change->Updates->SetItemField[] = $field;
	}
	$request->ItemChanges[] = $change;


	$response = $this->client->UpdateItem($request);

	// Iterate over the results, printing any error messages or ids of events that
	// were updated.
	$response_messages = $response->ResponseMessages->UpdateItemResponseMessage;
	foreach ($response_messages as $response_message) {
		// Make sure the request succeeded.
		if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
			$code = $response_message->ResponseCode;
			$message = $response_message->MessageText;
			fwrite(STDERR, "Failed to update event with \"$code: $message\"\n");
			continue;
	    	}
    		// Iterate over the updated events, printing the id of each.
    		foreach ($response_message->Items->CalendarItem as $item) {
        		$id = $item->ItemId->Id;
    		}
	}
    }

    public function deleteEvent($eventID, $changekey) {
	$request = new CreateItemType();
	$request->MessageDisposition = MessageDispositionType::SAVE_ONLY;
	$request->Items = new NonEmptyArrayOfAllItemsType();

	$cancellation = new CancelCalendarItemType();
	$cancellation->ReferenceItemId = new ItemIdType();
	$cancellation->ReferenceItemId->Id = $eventID;
	$cancellation->ReferenceItemId->ChangeKey = $changekey;
	$request->Items->CancelCalendarItem[] = $cancellation;

	$response = $this->client->CreateItem($request);

	// Iterate over the results, printing any error messages.
	$response_messages = $response->ResponseMessages->CreateItemResponseMessage;
	foreach ($response_messages as $response_message) {
		// Make sure the request succeeded.
		if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
			$code = $response_message->ResponseCode;
			$message = $response_message->MessageText;
			fwrite(STDERR, "Cancellation failed to create with \"$code: $message\"\n");
			continue;
    		}
	}
    }

    private function getChangeItemField($fieldName, $value, $type) {
	$field = new SetItemFieldType();
	$field->FieldURI = new PathToUnindexedFieldType();
	$field->FieldURI->FieldURI = $type;
	$field->CalendarItem = new CalendarItemType();
    	$field->CalendarItem->$fieldName = $value;
	return $field;
    }

}
