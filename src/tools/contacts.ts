import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface Contact {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: Array<{ address: string; name?: string }>;
  businessPhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  jobTitle?: string;
}

// Schemas
export const listContactsSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of contacts. Default: 50'),
});

export const searchContactsSchema = z.object({
  query: z.string().describe('Search query (name, email, or company)'),
  maxItems: z.number().optional().describe('Maximum results. Default: 25'),
});

export const getContactSchema = z.object({
  contactId: z.string().describe('The ID of the contact to retrieve'),
});

// Tool implementations
export async function listContacts(params: z.infer<typeof listContactsSchema>) {
  const { maxItems = 50 } = params;

  const path = `/me/contacts?$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle&$orderby=displayName&$top=${maxItems}`;

  const contacts = await graphList<Contact>(path, { maxItems });

  return contacts.map((c) => ({
    id: c.id,
    name: c.displayName,
    firstName: c.givenName,
    lastName: c.surname,
    email: c.emailAddresses?.[0]?.address,
    phone: c.mobilePhone || c.businessPhones?.[0],
    company: c.companyName,
    title: c.jobTitle,
  }));
}

export async function searchContacts(params: z.infer<typeof searchContactsSchema>) {
  const { query, maxItems = 25 } = params;

  // Use $filter for search since contacts don't support $search
  const encodedQuery = encodeURIComponent(query);
  const path = `/me/contacts?$filter=contains(displayName,'${encodedQuery}') or contains(givenName,'${encodedQuery}') or contains(surname,'${encodedQuery}')&$select=id,displayName,emailAddresses,businessPhones,mobilePhone,companyName&$top=${maxItems}`;

  const contacts = await graphList<Contact>(path, { maxItems });

  return contacts.map((c) => ({
    id: c.id,
    name: c.displayName,
    email: c.emailAddresses?.[0]?.address,
    phone: c.mobilePhone || c.businessPhones?.[0],
    company: c.companyName,
  }));
}

export async function getContact(params: z.infer<typeof getContactSchema>) {
  const { contactId } = params;

  const contact = await graphRequest<Contact>(
    `/me/contacts/${contactId}?$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle`
  );

  return {
    id: contact.id,
    name: contact.displayName,
    firstName: contact.givenName,
    lastName: contact.surname,
    emails: contact.emailAddresses?.map((e) => e.address),
    businessPhones: contact.businessPhones,
    mobilePhone: contact.mobilePhone,
    company: contact.companyName,
    title: contact.jobTitle,
  };
}
