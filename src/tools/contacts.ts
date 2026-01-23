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
  personalNotes?: string;
  birthday?: string;
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
  notes: z.string().optional().describe('Personal notes about the contact'),
  birthday: z.string().optional().describe('Birthday (ISO date format, e.g., 1990-05-15)'),
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
  notes: z.string().optional().describe('Personal notes about the contact'),
  birthday: z.string().optional().describe('Birthday (ISO date format, e.g., 1990-05-15)'),
});

// Tool implementations
export async function listContacts(params: z.infer<typeof listContactsSchema>) {
  const { maxItems = 50 } = params;

  const path = `/me/contacts?$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle,personalNotes,birthday&$orderby=displayName&$top=${maxItems}`;

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
    notes: c.personalNotes,
    birthday: c.birthday,
  }));
}

export async function searchContacts(params: z.infer<typeof searchContactsSchema>) {
  const { query, maxItems = 25 } = params;
  const queryLower = query.toLowerCase();

  // First, search using API filter (name, company, job title - Graph API supports these)
  const encodedQuery = encodeURIComponent(query);
  const filters = [
    `contains(displayName,'${encodedQuery}')`,
    `contains(givenName,'${encodedQuery}')`,
    `contains(surname,'${encodedQuery}')`,
    `contains(companyName,'${encodedQuery}')`,
    `contains(jobTitle,'${encodedQuery}')`,
  ].join(' or ');

  const path = `/me/contacts?$filter=${filters}&$select=id,displayName,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle,personalNotes,birthday&$top=${maxItems}`;

  let contacts: Contact[] = [];
  try {
    contacts = await graphList<Contact>(path, { maxItems });
  } catch {
    // If filter fails, fall back to getting all and filtering locally
  }

  // Also get all contacts to search notes and emails (not supported by $filter)
  const allPath = `/me/contacts?$select=id,displayName,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle,personalNotes,birthday&$top=500`;
  const allContacts = await graphList<Contact>(allPath, { maxItems: 500 });

  // Search locally in notes and emails
  const localMatches = allContacts.filter((c) => {
    const searchable = [
      c.personalNotes,
      ...(c.emailAddresses?.map((e) => e.address) || []),
    ].filter(Boolean);
    return searchable.some((s) => s?.toLowerCase().includes(queryLower));
  });

  // Merge results, dedupe by id
  const seen = new Set<string>();
  const merged: Contact[] = [];
  for (const c of [...contacts, ...localMatches]) {
    if (!seen.has(c.id)) {
      seen.add(c.id);
      merged.push(c);
    }
  }

  return merged.slice(0, maxItems).map((c) => ({
    id: c.id,
    name: c.displayName,
    email: c.emailAddresses?.[0]?.address,
    phone: c.mobilePhone || c.businessPhones?.[0],
    company: c.companyName,
    title: c.jobTitle,
    notes: c.personalNotes,
    birthday: c.birthday,
  }));
}

export async function getContact(params: z.infer<typeof getContactSchema>) {
  const { contactId } = params;

  const contact = await graphRequest<Contact>(
    `/me/contacts/${contactId}?$select=id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle,personalNotes,birthday`
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
    notes: contact.personalNotes,
    birthday: contact.birthday,
  };
}

export async function createContact(params: z.infer<typeof createContactSchema>) {
  const { givenName, surname, email, mobilePhone, businessPhone, companyName, jobTitle, notes, birthday } = params;

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
  if (notes) body.personalNotes = notes;
  if (birthday) body.birthday = birthday;

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
  const { contactId, givenName, surname, email, mobilePhone, businessPhone, companyName, jobTitle, notes, birthday } = params;

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
  if (notes !== undefined) body.personalNotes = notes;
  if (birthday !== undefined) body.birthday = birthday;

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
