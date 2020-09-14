#!/usr/bin/env bash

cd ~/hacking/awards || exit
wp term list award_category --field=term_id | xargs wp term delete award_category
wp term list award_year --field=term_id | xargs wp term delete award_year
wp post delete $(wp post list --post_type='award' --format=ids) --force
