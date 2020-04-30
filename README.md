# Approval Status

Get the status of the Approval flow that associated with a CDS entity record

I have built this control because of the need to let the user check the approval pending with whom, specially as there is no out of the box way to do it

In order to be able to use this control, Once you are building your approval flow, the item link should be matching with the record url (contain record id), as this is the only way to match the approval with the record (So far there is no regarding in the approval)

The user should have read privileges on these 3 entities (Approval, Approval Responses, Approval Requests) other wise, the data will not be loaded
