const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const csv = require('csv-parser');
const axios = require('axios');
const cors = require('cors')

const app = express();
const port = process.env.PORT || 3000;

// Multer storage configuration
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = 'uploads';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}`);
    }
});

// File filter for CSV only
const fileFilter = (req, file, cb) => {
    if (file.mimetype === 'text/csv') {
        cb(null, true);
    } else {
        cb(new Error('Only CSV files are allowed'), false);
    }
};

const upload = multer({
    storage,
    fileFilter,
    limits: { fileSize: 1024 * 1024 * 5 } // 5 MB limit
});

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use((cors()))
// CSV Upload Route
app.post('/api/csvfileupload', upload.single('csvFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'No file uploaded' });
        }
        const { url, cookie } = req.body;
        // console.log(url,cookie)
        try {

            if (!url || !cookie) {
                throw new Error("Missing 'url' or 'cookie' in ConnectionDetails");
            }
        } catch (err) {
            return res.status(400).json({ success: false, message: `Invalid ConnectionDetails: ${err.message}` });
        }

        const results = [];
        fs.createReadStream(req.file.path)
            .pipe(csv())
            .on('data', (data) => results.push(data))
            .on('end', async () => {
                try {
                    console.log(results, "results");
                    const processOutcome = await processSharedSpace(results, url, cookie);
                    res.status(200).json({
                        success: true,
                        message: 'CSV file uploaded and processed successfully',
                        file: {
                            filename: req.file.filename,
                            path: req.file.path,
                            size: req.file.size
                        },
                        processedData: processOutcome
                    });

                } catch (err) {
                    console.error('processSharedSpace error:', err.message);
                    res.status(500).json({ success: false, message: 'Error processing shared space', details: err.message });
                }
            })
            .on('error', (error) => {
                console.error('CSV parsing error:', error.message);
                res.status(500).json({ success: false, message: `Error parsing CSV: ${error.message}` });
            });

    } catch (err) {
        console.error('Upload handler error:', err.message);
        res.status(500).json({ success: false, message: err.message });
    }
});

// Get Widget ID
async function getWidgetId(url, cookie, companyId) {
    try {
        const response = await axios.post(`${url}/v2/galaxy/spaces/assignment/resolve/cid`, {
            companyId,
            entityId: companyId,
            entityType: "Company",
            sharingType: "external"
        }, {
            headers: {
                'Cookie': cookie,
                'Content-Type': 'application/json'
            },
            maxBodyLength: Infinity
        });

        const layout = response.data?.data?.layout;
        const section = layout?.sections?.[0].sectionId;
        const widget = layout?.sections?.[0]?.config?.widgets?.[0];
        console.log(layout, "getwidget");

        if (!layout || !section || !widget) {
            throw new Error('Incomplete layout or widget information');
        }

        return {
            layoutId: layout.layoutId,
            sectionId: section,
            widgetDetails: widget
        };
    } catch (error) {
        console.error(`getWidgetId failed for companyId=${companyId}:`, error);
        throw error;
    }
}


// Add Notes
async function addSpaceNotes(url, cookie, notes, companyId, sectionId, layoutId) {
    try {
        const data = {
            entityId: companyId,
            entityType: "company",
            layoutId,
            sectionId,
            data: {
                "SpacesNotes": notes
            }
        };
        console.log(data, "data")
        // https://demo-spaces.gainsightcloud.comv2/galaxy/spaces/cr360/data/section?ignoreMasked=true url

        console.log(url + "/v2/galaxy/spaces/cr360/data/section?ignoreMasked=true", "url")
        const response = await axios.put(url + "/v2/galaxy/spaces/cr360/data/section?ignoreMasked=true", data, {
            headers: {
                'Cookie': cookie,
                'Content-Type': 'application/json'
            },
            maxBodyLength: Infinity
        });

        console.log(`Notes updated for companyId=${companyId}`);
        return response.data;
    } catch (error) {
        console.error(`addSpaceNotes failed for companyId=${companyId}:`, error.message);
        throw error;
    }
}
async function addSucessPlan(url, cookie, Success_Plan_GSID) {
    try {
        const data = {
            "request": [
                {
                    "assetId": Success_Plan_GSID,
                    "sharedWithSpaces": true
                }
            ]
        }
        console.log(data, "data")
        const response = await axios.put(url + "/v1/successPlan/spaces-config", data, {
            headers: {
                'Cookie': cookie,
                'Content-Type': 'application/json'
            },
            maxBodyLength: Infinity
        });

        console.log(response.data, "success");
        return response.data;
    } catch (error) {
        console.error(`addSpaceNotes failed for companyId=${companyId}:`, error.message);
        throw error;
    }
}
async function trySearchUser(url, cookie, companyId, email) {
    try {
        const response = await axios.post(`${url}/v1/spaces/search/Company/person`, {
            entityId: companyId,
            companyId,
            searchString: email
        }, {
            headers: { 'Cookie': cookie, 'Content-Type': 'application/json' },
            maxBodyLength: Infinity
        });

        const person = response?.data?.data?.[0];
        if (person && person.person__Email && person.person__Gsid) {
            return person;
        } else {
            return null;
        }
    } catch (err) {
        console.error(`Failed to search user (${email}) for companyId=${companyId}:`, err.message);
        return null;
    }
}

async function sendInvitation(url, cookie, companyId, personId, email) {
    try {
        const data = {
            entityId: companyId,
            companyId,
            users: [{
                personId,
                email,
                userName: "System",  // Optional: customize
                permissionType: "NON_DELEGATE"
            }],
            emailBody: "",
            emailSubject: "Welcome to Spaces  ${subj::entity_name} "
        }

        const response = await axios.post(`${url}/v1/spaces/invite/company/persons`, data, {
            headers: { 'Cookie': cookie, 'Content-Type': 'application/json' },
            maxBodyLength: Infinity
        });

        console.log(`✅ Invitation sent to ${email}`);
        return response.data;
    } catch (error) {
        console.error(`❌ Invitation failed for ${email}:`, error.message);
        throw error;
    }
}
async function addUser(url, cookie, companyId, email) {
    try {
        const data = {
            "Name": email.split("@")[0],
            "Email": email,
            "companies": [
                {
                    "Company_ID": companyId
                }
            ]
        }

        const response = await axios.put(`${url}/v1/peoplemgmt/v1.0/people?areaName=PersonC360UI`, data, {
            headers: { 'Cookie': cookie, 'Content-Type': 'application/json' },
            maxBodyLength: Infinity
        });

        console.log(`✅ Invitation sent to ${email}`);
        return response.data;
    } catch (error) {
        console.error(`❌ Invitation failed for ${email}:`, error.message);
        throw error;
    }
}
async function updateWidgetDetails(url, cookie, companyGsid, layoutId, widgetDetails) {
    const apiUrl = `${url}/v2/galaxy/spaces/customisation/save/Company/${companyGsid}/${layoutId}`

    try {
        console.log("Yuva")
        console.dir(widgetDetails)
        const response = await fetch(apiUrl, {
            method: "PUT",
            headers: {
                Cookie: cookie,
                "Content-Type": "application/json",
            },
            body: JSON.stringify(widgetDetails),
        })

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`)
        }

        const result = await response.json()
        console.log("Widget updated successfully:", result)
        return result
    } catch (error) {
        console.error("Error updating widget:", error)
        throw error
    }
}
async function processSharedSpace(results, url, cookie) {
    const outcome = [];

    for (const row of results) {
        const { Company_GSID, Video_URL, Welcome_Banner, Space_Notes, Success_Plan_GSID, Invite_Email } = row;
        const recordResult = {
            Company_GSID,
            Video_URL,
            Invite_Email,
            status: "Success",
            messages: []
        };

        try {
            const { layoutId, sectionId, widgetDetails } = await getWidgetId(url, cookie, Company_GSID);
            console.log(widgetDetails, "widgetDetails");
            recordResult.messages.push("Fetched widget details");

            // Optional widget config
            if (widgetDetails?.config) {
                widgetDetails.config.mediaContent = {
                    mediaType: "VIDEO",
                    content: {
                        url: Video_URL || "",
                        thumbnailUrl: ""
                    }
                };

                if (Welcome_Banner) {
                    widgetDetails.config.bannerText = { value: Welcome_Banner };
                }
            }

            await updateWidgetDetails(url, cookie, Company_GSID, layoutId, widgetDetails);
            recordResult.messages.push("Widget updated");

            await addSpaceNotes(url, cookie, Space_Notes || "", Company_GSID, sectionId, layoutId);
            recordResult.messages.push("Notes added");

            await addSucessPlan(url, cookie, Success_Plan_GSID);
            recordResult.messages.push("Success Plan added");

            // Process each email individually
            const emailList = Invite_Email ? Invite_Email.split(',').map(e => e.trim()).filter(Boolean) : [];
            for (const email of emailList) {
                try {
                    let user = await trySearchUser(url, cookie, Company_GSID, email);
                    if (user) {
                        await sendInvitation(url, cookie, Company_GSID, user.person__Gsid, email);
                        recordResult.messages.push(`Invitation sent to ${email}`);
                    } else {
                        await addUser(url, cookie, Company_GSID, email);
                        user = await trySearchUser(url, cookie, Company_GSID, email);
                        if (user?.person__Gsid) {
                            await sendInvitation(url, cookie, Company_GSID, user.person__Gsid, email);
                            recordResult.messages.push(`User added and invitation sent to ${email}`);
                        } else {
                            recordResult.status = "Partial";
                            recordResult.messages.push(`User could not be added or invited for ${email}`);
                        }
                    }
                } catch (emailErr) {
                    recordResult.status = "Partial";
                    recordResult.messages.push(`Error with ${email}: ${emailErr.message}`);
                    console.error(`❌ Error processing invitation for ${email}:`, emailErr.message);
                }
            }

        } catch (err) {
            recordResult.status = "Failed";
            recordResult.messages.push(`Error: ${err.message}`);
            console.error(`❌ Error processing Company_GSID=${Company_GSID}:`, err.message);
        }

        outcome.push(recordResult);
    }

    return outcome;
}


app.use(express.static(path.join(__dirname, 'build')));
app.get('', (req, res) => {
    res.sendFile(path.join(__dirname, 'build', 'index.html'));
});


// Global Error Middleware
app.use((err, req, res, next) => {
    console.error('Unhandled error:', err.stack);
    res.status(500).json({ success: false, message: 'Internal Server Error', details: err.message });
});

// Start the server
app.listen(port, () => {
    console.log(`🚀 Server is running on port ${port}`);
});
