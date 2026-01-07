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

export const createContactSchema = z.object({
  givenName: z.string().optional().describe('First name'),
  surname: z.string().optional().describe('Last name'),
  email: z.string().optional().describe('Email address'),
  mobilePhone: z.string().optional().describe('Mobile phone number'),
  businessPhone: z.string().optional().describe('Business phone number'),
  companyName: z.string().optional().describe('Company name'),
  jobTitle: z.string().optional().describe('Job title'),
});

export const updateContactSchema = z.object({
  contactId: z.string().describe('The ID of the contact to update'),
  givenName: z.string().optional().describe('First name'),
  surname: z.string().optional().describe('Last name'),
  email: z.string().optional().describe('Email address'),
  mobilePhone: z.string().optional().describe('Mobile phone number'),
  businessPhone: z.string().optional().describe('Business phone number'),
  companyName: z.string().optional().describe('Company name'),
  jobTitle: z.string().optional().describe('Job title'),
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

export async function createContact(params: z.infer<typeof createContactSchema>) {
  const { givenName, surname, email, mobilePhone, businessPhone, companyName, jobTitle } = params;

  const body: Record<string, unknown> = {};

  if (givenName) body.givenName = givenName;
  if (surname) body.surname = surname;
  if (email) {
    body.emailAddresses = [{ address: email, name: `${givenName || ''} ${surname || ''}`.trim() || email }];
  }
  if (mobilePhone) body.mobilePhone = mobilePhone;
  if (businessPhone) body.businessPhones = [businessPhone];
  if (companyName) body.companyName = companyName;
  if (jobTitle) body.jobTitle = jobTitle;

  const contact = await graphRequest<Contact>('/me/contacts', {
    method: 'POST',
    body,
  });

  return {
    success: true,
    contactId: contact.id,
    name: contact.displayName,
    message: 'Contact created',
  };
}

export async function updateContact(params: z.infer<typeof updateContactSchema>) {
  const { contactId, givenName, surname, email, mobilePhone, businessPhone, companyName, jobTitle } = params;

  const body: Record<string, unknown> = {};

  if (givenName !== undefined) body.givenName = givenName;
  if (surname !== undefined) body.surname = surname;
  if (email !== undefined) {
    body.emailAddresses = [{ address: email }];
  }
  if (mobilePhone !== undefined) body.mobilePhone = mobilePhone;
  if (businessPhone !== undefined) body.businessPhones = [businessPhone];
  if (companyName !== undefined) body.companyName = companyName;
  if (jobTitle !== undefined) body.jobTitle = jobTitle;

  const contact = await graphRequest<Contact>(`/me/contacts/${contactId}`, {
    method: 'PATCH',
    body,
  });

  return {
    success: true,
    contactId: contact.id,
    name: contact.displayName,
    message: 'Contact updated',
  };
}
